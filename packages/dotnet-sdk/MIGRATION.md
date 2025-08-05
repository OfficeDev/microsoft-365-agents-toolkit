# TeamsFx .NET SDK

## Notification Bot

TeamsFx SDK provides `ConversationBot.Notification` to send proactive notification message to Bot installations, such as personal installations, channels and groupChats.

The main thing to note is listen to `ConversationUpdateEvents.MembersAdded` event, store `ConversationReference` somewhere, then you can use it to send proactive notification later.

With Microsoft.TeamsFx SDK, you can simply create a ConversationBot:

```csharp
// In Program.cs
builder.Services.AddSingleton(sp =>
{
    var options = new ConversationOptions()
    {
        Adapter = sp.GetService<CloudAdapter>(),
        Notification = new NotificationOptions
        {
            BotAppId = builder.Configuration["Connections:BotServiceConnection:Settings:ClientId"],
        },
    };

    return new ConversationBot(options);
});

// In Trigger.cs
var pagedInstallations = await _conversation.Notification.GetPagedInstallationsAsync(pageSize, continuationToken, req.HttpContext.RequestAborted);
```

Without TeamsFx SDK, you can put key classes `NotificationBot.cs`, `NotificationMiddleware.cs`, and other notification-related classes into your Teams app source code, use `LocalFileStorage` or implement your own persistent storage.

```csharp
builder.Services.AddSingleton(sp =>
{
    var adapter = sp.GetService<CloudAdapter>();
    var botAppId = builder.Configuration["Connections:BotServiceConnection:Settings:ClientId"];
    return new NotificationBot(adapter, botAppId, null);
});

// In Program.cs
builder.Services.AddSingleton(sp =>
{
    var options = new ConversationOptions()
    {
        Adapter = sp.GetService<CloudAdapter>(),
        Notification = new NotificationOptions
        {
            BotAppId = builder.Configuration["Connections:BotServiceConnection:Settings:ClientId"],
        },
    };

    return new ConversationBot(options);
});

// In Trigger.cs
var pagedInstallations = await _notification.GetPagedInstallationsAsync(pageSize, continuationToken, cancellationToken);
```

You can refer to [notification-http-trigger template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vs/csharp/notification-http-trigger).

## Command Bot and Workflow Bot

With TeamsFx SDK, you can define a ConversationBot with commands.

```csharp
// In Program.cs
builder.Services.AddSingleton<HelloWorldCommandHandler>();

builder.Services.AddSingleton(sp =>
{
    var options = new ConversationOptions()
    {
        Adapter = sp.GetService<CloudAdapter>(),
        Command = new CommandOptions()
        {
            Commands = new List<ITeamsCommandHandler> { 
                sp.GetService<HelloWorldCommandHandler>()
            }
        },
        CardAction = new CardActionOptions()
        {
            Actions = new List<IAdaptiveCardActionHandler> { sp.GetService<DoStuffActionHandler>() }
        }
    };
    return new ConversationBot(options);
});

// Command handler implementation
public class HelloWorldCommandHandler : ITeamsCommandHandler
{
    public IEnumerable<ITriggerPattern> TriggerPatterns => new List<ITriggerPattern>
    {
        new RegExpTrigger("helloworld")
    };

    public async Task<ICommandResponse> HandleCommandAsync(ITurnContext context, CommandMessage message, CancellationToken cancellationToken = default)
    {
        // Your logic here
    }
}
```

### Option: Use Microsoft Agents SDK

[Microsoft Agents SDK](https://www.nuget.org/packages/Microsoft.Agents.Hosting.AspNetCore) provides a simple set of functions over the Microsoft Bot Framework to implement this scenario.

```csharp
// In TeamsBot.cs
public class TeamsBot : AgentApplication
{
    public TeamsBot(AgentApplicationOptions options) : base(options)
    {
        OnConversationUpdate(ConversationUpdateEvents.MembersAdded, OnMembersAddedAsync);
        OnActivity(ActivityTypes.Message, OnMessageReceivedAsync);
        AdaptiveCards.OnActionExecute(DoStuffActionHandler.TriggerVerb, DoStuffActionHandler.handler);
    }

    protected async Task OnMessageReceivedAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        var messageText = turnContext.Activity.Text?.Trim().ToLowerInvariant();
        
        if (messageText == "helloworld")
        {
            await turnContext.SendActivityAsync(MessageFactory.Text("Hello World!"), cancellationToken);
        }
        else if (messageText == "help")
        {
            await turnContext.SendActivityAsync(MessageFactory.Text("Available commands: helloworld, help"), cancellationToken);
        }
    }
}

// In DoStuffActionHandler.cs
public class DoStuffActionHandler
    {
        public static string TriggerVerb => "doStuff";

        public static ActionExecuteHandler handler = HandleActionInvokedAsync;

        private static async Task<AdaptiveCardInvokeResponse> HandleActionInvokedAsync(ITurnContext turnContext, ITurnState state, object cardData, CancellationToken cancellationToken = default)
        {
            // Render adaptive card content
            var cardContent = new AdaptiveCardTemplate(cardTemplate).Expand
            (
                new HelloWorldModel
                {
                    Title = "Hello World Bot",
                    Body = $"Congratulations! Your {TriggerVerb} action is processed successfully!",
                }
            );

            // Send invoke response with adaptive card
            return  AdaptiveCardInvokeResponseFactory.AdaptiveCard(cardContent);
        }
    }
```

You can refer to [command and response template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vs/csharp/command-and-response) and [workflow template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vs/csharp/workflow).

## Bot SSO

### Option: Use Teams AI Library

[Teams AI Library](https://www.nuget.org/packages/Microsoft.Teams.AI) integrates with `TeamsBotSsoPrompt`. You can add authentication configurations to Application.

```csharp
// Program.cs
    IConfidentialClientApplication msal = sp.GetService<IConfidentialClientApplication>();
    string signInLink = $"https://{config.BOT_DOMAIN}/auth-start.html";

    AuthenticationOptions<AppState> options = new();
    options.AddAuthentication("graph", new TeamsSsoSettings(new string[]{"User.Read"}, signInLink, msal));

    Application<AppState> app = new ApplicationBuilder<AppState>()
        .WithStorage(storage)
        .WithTurnStateFactory(() => new AppState())
        .WithAuthentication(adapter, options)
        .Build();


    app.Authentication.Get("graph").OnUserSignInSuccess(async (context, state) =>
    {
        await context.SendActivityAsync("Successfully logged in");
        // Get token from state
        await context.SendActivityAsync($"Token string length: {state.Temp.AuthTokens["graph"].Length}");
        await context.SendActivityAsync($"This is what you said before the AuthFlow started: {context.Activity.Text}");
    });

```

You can refer to [Teams AI bot sso sample](https://github.com/microsoft/teams-ai/blob/main/dotnet/samples/06.auth.teamsSSO.bot)

## Tab SSO

### Option 1: Move TeamsUserCredential.cs into source code

`TeamsUserCredential` represents Teams current user's identity. Using this credential will request user consent at the first time. It leverages the Teams SSO and On-Behalf-Of flow to do token exchange. SDK uses this credential when developer choose "User" identity in browser environment. You can copy this `TeamsUserCredential.cs` into your source code.

### Option 2: NAA

We recommend [Nested App Auth](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/authentication/nested-authentication) to implement SSO.

You can refer to [NAA sample](https://github.com/OfficeDev/Microsoft-Teams-Samples/tree/main/samples/tab-nested-auth/csharp).
