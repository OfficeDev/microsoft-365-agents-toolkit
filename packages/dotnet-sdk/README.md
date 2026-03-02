# TeamsFx .NET SDK

This package will be deprecated by 2026-09.

A NuGet package for Blazor projects which aims to reduce the developer tasks of implementing identity and access to cloud resources.

[TeamsFx SDK for JavaScript/TypeScript](https://github.com/OfficeDev/TeamsFx/tree/main/packages/sdk) |
[API reference documentation](https://aka.ms/teamsfx-sdk-help)

## Getting started

Build Teams apps with Blazor and the TeamsFx .NET SDK using Teams Toolkit. [Visit the documentation to learn more](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/visual-studio-overview).

### Prerequisites

1. Install the `ASP.NET and web development` workload using the Visual Studio Installer.
2. Launch Visual Studio and create a new Blazor project.

## Usage

### How to get the package

1. Right-click on the project in Visual Studio and choose Manage NuGet Packages.
2. Search for `Microsoft.TeamsFx` and add it to the Blazor project.

Alternately, you can use the Package Manager.

```ps
> Install-Package Microsoft.TeamsFx
```

### How to choose version

For .NET 5 projects (VS 2019): Choose version < 0.3.0-rc.
For .NET 6 projects (VS 2022): Choose version >= 0.3.0-rc.

### Using Teams User Credential in Teams Tab app

1. Add authentication options in appsettings.{Environment}.json file.

```json
"TeamsFx": {
    "Authentication": {
        "ClientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "ClientSecret": "xxx", // 'User Secrets' is a better place to store secret string.
        "OAuthAuthority": "https://login.microsoftonline.com/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    }
}
```

2. Add `TeamsFx` to services during startup.

```csharp
public void ConfigureServices(IServiceCollection services)
{
    ...
    services.AddTeamsFx(Configuration.GetSection("TeamsFx"));
}
```

3. Add the required namespaces to the `_Imports.razor` file.

```csharp
@using Microsoft.TeamsFx
```

4. Inject the registered TeamsFx services for any page that needs them.

```csharp
@inject TeamsFx teamsfx
@inject TeamsUserCredential teamsUserCredential
```

5. Call `teamsUserCredential.GetTokenAsync()` to get access token or pass `teamsUserCredential` to other functions.

```csharp
try
{
    await teamsUserCredential.GetTokenAsync(
        new TokenRequestContext(new string[] { "User.Read" }),
        new System.Threading.CancellationToken());
}
catch (ExceptionWithCode e)
{
    if (e.Code == ExceptionCode.UiRequiredError)
    {
        // show login button to let user consent
    }
    else
    {
        throw;
    }
}
```

### Using Graph Service with Teams User Credential in Teams Tab app

1. Follow the above `Using Teams User Credential in Teams Tab app` to create `TeamsUserCredential`.
2. Add `@using Microsoft.Graph` to any page that needs the Graph service.
3. Initialize a Graph Service Client with credential and scopes.

   ```csharp
   string _scope = "User.Read";
   var client = new GraphServiceClient(teamsUserCredential, new string[] { _scope });
   ```

4. To make requests against the service, you could follow the [guide](https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/docs/overview.md#requests). For instance, if you wish to get user profile, you could call `graph.Me.GetAsync();`

### Using Conversation Bot for Command and Response

1. Add your command handler class which implements the `ITeamsCommandHandler` interface.

   - Define your trigger patters in the `TriggerPatterns` property, you can use `StringTrigger` or `RegExpTrigger`.
   - Handle your command in `HandleCommandAsync` function, and return an `ActivityCommandResponse` or `TextommandResponse` object as the command response.

   ```csharp
   public class SampleCommandHandler : ITeamsCommandHandler
   {
       // Define your trigger patterns
       public IEnumerable<ITriggerPattern> TriggerPatterns => new List<ITriggerPattern>
       {
           new StringTrigger("helloworld")
       };

       // Handle your command and send response to Teams chat
       public async Task<ICommandResponse> HandleCommandAsync(ITurnContext turnContext, CommandMessage message, CancellationToken cancellationToken = default)
       {
           // TODO: provide your implementation here.
           return new TextCommandResponse("This is a sample response!");
       }
   }
   ```

2. Initialize the command bot and register your commands in your app's startup (usually it's in `Program.cs` or `Startup.cs`)

   ```csharp
   builder.Services.AddSingleton<SampleCommandHandler>();
   builder.Services.AddSingleton(sp =>
   {
       var options = new ConversationOptions()
       {
           // NOTE: you need to register your CloudAdapter into your service before conversation bot initialization.
           Adapter = sp.GetService<CloudAdapter>(),
           Command = new CommandOptions()
           {
               Commands = new List<ITeamsCommandHandler> { sp.GetService<SampleCommandHandler>() }
           }
       };

       return new ConversationBot(options);
   });
   ```

3. Use the conversation bot in your bot controller

   ```csharp
   namespace SampleTeamsApp.Controllers
   {
       using Microsoft.AspNetCore.Mvc;
       using Microsoft.Bot.Builder;
       using Microsoft.Bot.Builder.Integration.AspNet.Core;
       using Microsoft.TeamsFx.Conversation;

       [Route("api/messages")]
       [ApiController]
       public class BotController : ControllerBase
       {
           private readonly ConversationBot _conversation;
           private readonly IBot _bot;

           public BotController(ConversationBot conversation, IBot bot)
           {
               _conversation = conversation;
               _bot = bot;
           }

           [HttpPost]
           public async Task PostAsync(CancellationToken cancellationToken = default)
           {
               await (_conversation.Adapter as CloudAdapter).ProcessAsync
               (
                   Request,
                   Response,
                   _bot,
                   cancellationToken
               );
           }
       }
   }
   ```

### Using Conversation Bot for Notification

1. Initialize your own bot adapter and the `ConversationBot` in your app's startup (usually it's in `Program.cs` or `Startup.cs`)

   ```csharp
   // Create the Conversation with notification feature enabled.
   builder.Services.AddSingleton(sp =>
   {
       var options = new ConversationOptions()
       {
           // NOTE: you need to register your CloudAdapter into your service before conversation bot initialization.
           Adapter = sp.GetService<CloudAdapter>(),
           Notification = new NotificationOptions
           {
               BotAppId = botAppId, // Your bot app ID
           },
       };

       return new ConversationBot(options);
   });
   ```

2. Reference the conversation bot in your bot message controller/handler to ensure it's initialized before handling any bot message

   ```csharp
   namespace SampleTeamsApp.Controllers
   {
       using Microsoft.AspNetCore.Mvc;
       using Microsoft.Bot.Builder;
       using Microsoft.Bot.Builder.Integration.AspNet.Core;
       using Microsoft.TeamsFx.Conversation;

       [Route("api/messages")]
       [ApiController]
       public class BotController : ControllerBase
       {
           private readonly ConversationBot _conversation;
           private readonly IBot _bot;

           public BotController(ConversationBot conversation, IBot bot)
           {
               _conversation = conversation;
               _bot = bot;
           }

           [HttpPost]
           public async Task PostAsync(CancellationToken cancellationToken = default)
           {
               await (_conversation.Adapter as CloudAdapter).ProcessAsync
               (
                   Request,
                   Response,
                   _bot,
                   cancellationToken
               );
           }
       }
   }
   ```

3. Send notification (called by your own controller or trigger)

   ```csharp
    public async Task NotifyAsync(ConversationBot conversation, CancellationToken cancellationToken)
    {
        var pageSize = 100;
        string continuationToken = null;
        do
        {
            var pagedInstallations = await conversation.Notification.GetPagedInstallationsAsync(pageSize, continuationToken, cancellationToken);
            continuationToken = pagedInstallations.ContinuationToken;
            var installations = pagedInstallations.Data;
            foreach (var installation in installations)
            {
                await installation.SendMessage("Hello.", cancellationToken);

                // Or, send adaptive card (need to build your own card object)
                // await installation.SendAdaptiveCard(cardObject, cancellationToken);
            }

        } while (!string.IsNullOrEmpty(continuationToken));
    }
   ```

### Using Conversation Bot for Adaptive Card Actions

1. Add your adaptive card action handler class which implements the `IAdaptiveCardActionHandler` interface.

   - Set the `TriggerVerb` property, the value should be the same as the `verb` property of the `Action.Execute` action.
   - Handle your action in `HandleActionInvokedAsync` function, and return an `InvokeResponse` as the action response.

   ```csharp
   public class DoStuffActionHandler : IAdaptiveCardActionHandler
   {
       /// <summary>
       /// A global unique string associated with the `Action.Execute` action.
       /// The value should be the same as the `verb` property which you define in your adaptive card JSON.
       /// </summary>
       public string TriggerVerb => "doStuff";

       /// <summary>
       /// Indicate how your acrion response card is sent in the conversation.
       /// By default, the response card can only be updated for the interactor who trigger the action.
       /// </summary>
       public AdaptiveCardResponse AdaptiveCardResponse => AdaptiveCardResponse.ReplaceForInteractor;


       public async Task<InvokeResponse> HandleActionInvokedAsync(ITurnContext turnContext, object cardData, CancellationToken cancellationToken = default)
       {
           // Send invoke response with text message
           return InvokeResponseFactory.TextMessage("[ACK] Successfully!");

           /**
            * If you want to send invoke response with adaptive card, you can:
            *
            * return InvokeResponseFactory.AdaptiveCard(JsonConvert.DeserializeObject(<your-card-json>));
            */

           /**
            * If you want to send invoke response with error message, you can:
            *
            * return InvokeResponseFactory.ErrorResponse(InvokeResponseErrorCode.BadRequest, "The incoming request is invalid.");
            */
       }
   }
   ```

2. Initialize your own bot adapter and the `ConversationBot` in your app's startup (usually it's in `Program.cs` or `Startup.cs`)

   ```csharp
   // create action handler instance
   builder.Services.AddSingleton<DoStuffActionHandler>();

   // create conversation bot with adaptive card action feature enabled.
   builder.Services.AddSingleton(sp =>
   {
       var options = new ConversationOptions()
       {
           // NOTE: you need to register your CloudAdapter into your service before conversation bot initialization.
           Adapter = sp.GetService<CloudAdapter>(),
           CardAction = new CardActionOptions()
           {
               Actions = new List<IAdaptiveCardActionHandler> { sp.GetService<DoStuffActionHandler>() }
           }
       };

       return new ConversationBot(options);
   });
   ```

3. Reference the conversation bot in your bot message controller/handler to ensure it's initialized before handling any bot message

   ```csharp
   namespace SampleTeamsApp.Controllers
   {
       using Microsoft.AspNetCore.Mvc;
       using Microsoft.Bot.Builder;
       using Microsoft.Bot.Builder.Integration.AspNet.Core;
       using Microsoft.TeamsFx.Conversation;

       [Route("api/messages")]
       [ApiController]
       public class BotController : ControllerBase
       {
           private readonly ConversationBot _conversation;
           private readonly IBot _bot;

           public BotController(ConversationBot conversation, IBot bot)
           {
               _conversation = conversation;
               _bot = bot;
           }

           [HttpPost]
           public async Task PostAsync(CancellationToken cancellationToken = default)
           {
               await (_conversation.Adapter as CloudAdapter).ProcessAsync
               (
                   Request,
                   Response,
                   _bot,
                   cancellationToken
               );
           }
       }
   }
   ```

## SDK Upgrade Steps

### Upgrade from 0.1.0-rc to 0.3.0 (For projects created by Visual Studio 2019 toolkit)

If there is an existing project created in VS2019, you can use the following steps to upgrade:

1. Open project in VS2022 and change project target framework to ".NET 6".

2. Upgrade dependencies:
   `Microsoft.TeamsFx.SimpleAuth` to `0.1.2`,
   `Newtonsoft.Json` to `13.0.1`,
   `Microsoft.Graph` to `4.12.0`,
   `Microsoft.Fast.Components.FluentUI` to `1.1.0`.

3. Add following lines in appsettings.{Environment}.json file after "ALLOWED_APP_IDS".

```json
"ALLOWED_APP_IDS": "...",
"TeamsFx": {
    "Authentication": {
        "ClientId": "value copied from CLIENT_ID",
        "SimpleAuthEndpoint": "value copied from TAB_APP_ENDPOINT",
        "InitiateLoginEndpoint": "{value copied from TAB_APP_ENDPOINT}/auth-start.html"
    }
}
```

4. Add following lines in `Startup.cs`.

```csharp
public void ConfigureServices(IServiceCollection services)
{
    ...
    services.AddTeamsFx(Configuration.GetSection("TeamsFx"));
}
```

and remove following 2 lines.

```csharp
services.AddScoped<TeamsFx>();
services.AddScoped<TeamsUserCredential>();
```

5. Remove following codes in `Welcome.razor`.

```csharp
var clientId = Configuration.GetValue<string>("CLIENT_ID");
var endpoint = MyNavigationManager.BaseUri;

await teamsfx.SetLogLevelAsync(LogLevel.Verbose);
await teamsfx.SetLogFunctionAsync(printLog);

AuthenticationConfiguration authentication = new AuthenticationConfiguration(clientId: clientId, simpleAuthEndpoint: endpoint, initiateLoginEndpoint: endpoint + "auth-start.html");
Configuration configuration = new Configuration(authentication);
await teamsfx.LoadConfigurationAsync(configuration);
...
private void printLog(LogLevel level, string message)
{
    Console.WriteLine(message);
}
```

### Upgrade from 0.3.0-rc to 0.4.0-rc (For projects created by Visual Studio 2022 17.1 Preview toolkit)

If there is an existing project created in VS2022 17.1 Preview, you can use the following steps to upgrade:

- In `appsettings.{Environment}.json` file:

1. Add `OAuthAuthority` under `TeamsFx:Authentication` and copy the value from `OAUTH_AUTHORITY`.
2. Remove the line `"SimpleAuthEndpoint": "https://localhost:port/"`.
3. Remove lines of configuration starting with "CLIENT_ID", "IDENTIFIER_URI", "TAB_APP_ENDPOINT", "OAUTH_AUTHORITY", "AAD_METADATA_ADDRESS", "ALLOWED_APP_IDS".
4. Remove the Nuget dependency package "Microsoft.TeamsFx.SimpleAuth".

- In Solution Explorer:

1. Right click project file and choose "Manage User Secrets".
2. Change key name "CLIENT_SECRET" to "TeamsFx:Authentication:ClientSecret".

### Upgrade from 1.1.0 to 1.2.0 (Update projects to use TeamsJS V.2.0)

Teams Toolkit provides users with a template `TeamsJSBlazorInterop.js`, which consists of multiple commonly used Teams JS SDK API. Users can add more APIs if needed. As suggested, even if you intend your app to only run in Teams (and not Outlook and the Microsoft 365 app), best practice is to start referencing the latest TeamsJS (_v.2.0_ or later) as soon as convenient, in order to benefit from the latest improvements, new features, and support (even for Teams-only apps).

Starting from TeamsFx .NET SDK 1.2.0, TeamsJS V2 are referenced. Though previously scaffolded projects still work with TeamsFx .NET SDK 1.2.0, we suggest you replace the `TeamsJSBlazorInterop.js` under `./wwwroot/js` with the latest one from [here](https://github.com/OfficeDev/TeamsFx/blob/main/templates/tab/csharp/default/wwwroot/js/TeamsJsBlazorInterop.js). Two APIs in the file are updated with new function name to align with renamed TeamsJS APIs. `initializeWithContext()` is removed.

| Original Function              | New Function                  |
| ------------------------------ | ----------------------------- |
| setFrameContext()              | setCurrentFrame()             |
| registerChangeSettingHandler() | registerChangeConfigHandler() |

For more APIs, please [visit the TeamsJS documentation to learn more](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-sdk?tabs=javascript%2Cmanifest-teams-toolkit).

### Upgrade from 1.x.x to 2.0.0

In this release, TeamsFx SDK add supports for Graph SDK v5 and remove `frameworkreference`. `auth-start.html` and `auth-end.html` are removed from SDK, and added to templates. You could use the following steps to upgrade existing projects.

1. Upgrade package dependencies in `{{ProjectName}}.csproj`.

   ```csharp
   <PackageReference Include="Microsoft.Graph" Version="5.6.0" />
   <PackageReference Include="Microsoft.TeamsFx" Version="2.0.0" />
   ```

2. If you are using class `User`, add `@using Microsoft.Graph.Models` to the top of the file.

3. Update requests to Microsoft Graph as follows.

   ```csharp
   // Previous
   var Profile = await graph.Me.Request().GetAsync();
   var photoStream = await graph.Me.Photo.Content.Request().GetAsync();

   // Now
   var Profile = await graph.Me.GetAsync();
   var photoStream = await graph.Me.Photo.Content.GetAsync();
   ```

4. Download `auth-start.html` and `auth-end.html` from [GitHub Repo](https://github.com/OfficeDev/TeamsFx/tree/dev/templates/csharp/sso-tab/wwwroot) to `{ProjectDirectory}/wwwroot`.

5. Update `appsetting.json` and `appsettings.Development.json` to add `InitiateLoginEndpoint`.

   ```json
   {
     "TeamsFx": {
       "Authentication": {
         "ClientId": "$clientId$",
         "ClientSecret": "$client-secret$",
         "InitiateLoginEndpoint": "$TAB_ENDPOINT$/auth-start.html", //New Line
         "OAuthAuthority": "$oauthAuthority$"
       }
     }
   }
   ```

6. Update `azure.bicep` under `{ProjectDirectory}/infra`.

   ```bicep
   // Previous
   resource webApp 'Microsoft.Web/sites@2021-02-01' = {
     kind: 'app'
     location: location
     name: webAppName
     properties: {
       serverFarmId: serverfarm.id
       httpsOnly: true
       siteConfig: {
         appSettings: [
           {
             name: 'WEBSITE_RUN_FROM_PACKAGE'
             value: '1'
           }
           {
             name: 'TeamsFx__Authentication__ClientId'
             value: tabAadAppClientId
           }
           {
             name: 'TeamsFx__Authentication__ClientSecret'
             value: tabAadAppClientSecret
           }
           {
             name: 'TeamsFx__Authentication__OAuthAuthority'
             value: uri(tabAadAppOauthAuthorityHost, tabAadAppTenantId)
           }
         ]
         ftpsState: 'FtpsOnly'
       }
     }
   }

   // Updated
   resource webApp 'Microsoft.Web/sites@2021-02-01' = {
     kind: 'app'
     location: location
     name: webAppName
     properties: {
       serverFarmId: serverfarm.id
       httpsOnly: true
       siteConfig: {
         ftpsState: 'FtpsOnly'
       }
     }
   }

   resource  webAppConfig  'Microsoft.Web/sites/config@2021-02-01' = {
     name: '${webAppName}/appsettings'
     properties: {
       WEBSITE_RUN_FROM_PACKAGE: '1'
       TeamsFx__Authentication__ClientId: tabAadAppClientId
       TeamsFx__Authentication__ClientSecret: tabAadAppClientSecret
       TeamsFx__Authentication__InitiateLoginEndpoint: 'https://${webApp.properties.defaultHostName}/auth-start.html'
       TeamsFx__Authentication__OAuthAuthority: uri(tabAadAppOauthAuthorityHost, tabAadAppTenantId)
     }
   }
   ```

7. [Optional] Update `teamsapp.local.yml` as follows.

   ```yaml
   - uses: file/updateJson # Generate runtime appsettings to JSON file
     with:
     target: ./appsettings.Development.json
     appsettings:
       TeamsFx:
         Authentication:
           ClientId: ${{AAD_APP_CLIENT_ID}}
           ClientSecret: ${{SECRET_AAD_APP_CLIENT_SECRET}}
           InitiateLoginEndpoint: ${{TAB_ENDPOINT}}/auth-start.html # New line
           OAuthAuthority: ${{AAD_APP_OAUTH_AUTHORITY}}
   ```

### Configure Logging

`ILogger` is used to print logs. You can configure logging in appsettings.{Environment}.json. [Visit the ASP.NET documentation to learn more](https://docs.microsoft.com/en-us/aspnet/core/fundamentals/logging/?view=aspnetcore-5.0#configure-logging-1)

### Bot

Will be supported in the future.

## Data Collection.

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at https://go.microsoft.com/fwlink/?LinkID=824704. You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)

If you are interested in fixing issues and contributing directly to the code base, please see the [Contributing Guide](./CONTRIBUTING.md).

## Reporting Security Issues

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them to the Microsoft Security Response Center (MSRC) at [https://msrc.microsoft.com/create-report](https://msrc.microsoft.com/create-report).

If you prefer to submit without logging in, send email to [secure@microsoft.com](mailto:secure@microsoft.com). If possible, encrypt your message with our PGP key; please download it from the the [Microsoft Security Response Center PGP Key page](https://www.microsoft.com/en-us/msrc/pgp-key-msrc).

You should receive a response within 24 hours. If for some reason you do not, please follow up via email to ensure we received your original message. Additional information can be found at [microsoft.com/msrc](https://www.microsoft.com/msrc).

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE.txt) license.
