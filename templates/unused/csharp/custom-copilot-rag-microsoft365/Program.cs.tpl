using {{SafeProjectName}};
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Teams.AI;
using Microsoft.Teams.AI.AI.Models;
using Microsoft.Teams.AI.AI.Planners;
using Microsoft.Teams.AI.AI.Prompts;
using Microsoft.Teams.AI.State;
using Microsoft.Teams.AI.AI;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Newtonsoft.Json.Linq;
using {{SafeProjectName}}.Model;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();

// Prepare Configuration for ConfigurationBotFrameworkAuthentication
var config = builder.Configuration.Get<ConfigOptions>();
builder.Configuration["MicrosoftAppType"] = config.BOT_TYPE;
builder.Configuration["MicrosoftAppId"] = config.BOT_ID;
builder.Configuration["MicrosoftAppPassword"] = config.BOT_PASSWORD;
builder.Configuration["MicrosoftAppTenantId"] = config.BOT_TENANT_ID;
// Create the Bot Framework Authentication to be used with the Bot Adapter.
builder.Services.AddSingleton<BotFrameworkAuthentication, ConfigurationBotFrameworkAuthentication>();

// Create the Cloud Adapter with error handling enabled.
// Note: some classes expect a BotAdapter and some expect a BotFrameworkHttpAdapter, so
// register the same adapter instance for both types.
builder.Services.AddSingleton<TeamsAdapter, AdapterWithErrorHandler>();
builder.Services.AddSingleton<IBotFrameworkHttpAdapter>(sp => sp.GetService<TeamsAdapter>());
builder.Services.AddSingleton<BotAdapter>(sp => sp.GetService<TeamsAdapter>());

builder.Services.AddSingleton<IStorage, MemoryStorage>();

builder.Services.AddSingleton<OpenAIModel>(sp => new(
{{#useOpenAI}}
    new OpenAIModelOptions(config.OpenAI.ApiKey, config.OpenAI.DefaultModel)
{{/useOpenAI}}
{{#useAzureOpenAI}}
    new AzureOpenAIModelOptions(
        config.Azure.OpenAIApiKey,
        config.Azure.OpenAIDeploymentName,
        config.Azure.OpenAIEndpoint
    )
{{/useAzureOpenAI}}
    {
        LogRequests = true,
    },
    sp.GetService<ILoggerFactory>()
));

builder.Services.AddSingleton(sp =>
{
    IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(config.AAD_APP_CLIENT_ID)
                                        .WithClientSecret(config.AAD_APP_CLIENT_SECRET)
                                        .WithTenantId(config.AAD_APP_TENANT_ID)
                                        .WithLegacyCacheCompatibility(false)
                                        .Build();
    app.AddInMemoryTokenCache(); // For development purpose only, use distributed cache in production environment
    return app;
});

GraphDataSource myDataSource = new GraphDataSource("graph-ai-search");

// Create the bot as transient. In this case the ASP Controller is expecting an IBot.
builder.Services.AddTransient<IBot>(sp =>
{
    // Create loggers
    ILoggerFactory loggerFactory = sp.GetService<ILoggerFactory>();

    // Create Prompt Manager
    PromptManager prompts = new(new()
    {
        PromptFolder = "./Prompts"
    });
    prompts.AddDataSource("graph-ai-search", myDataSource);

    // Create ActionPlanner
    ActionPlanner<AppState> planner = new(
        options: new(
            model: sp.GetService<OpenAIModel>(),
            prompts: prompts,
            defaultPrompt: async (context, state, planner) =>
            {
                PromptTemplate template = prompts.GetPrompt("chat");
                return await Task.FromResult(template);
            }
        )
        { LogRepairs = true },
        loggerFactory: loggerFactory
    );

    AIOptions<AppState> options = new(planner);
    options.EnableFeedbackLoop = true;

    IStorage storage = sp.GetService<IStorage>()!;
    TeamsAdapter adapter = sp.GetService<TeamsAdapter>()!;
    IConfidentialClientApplication msal = sp.GetService<IConfidentialClientApplication>();
    string signInLink = $"https://{config.BOT_DOMAIN}/auth-start.html";
    AuthenticationOptions<AppState> authOptions = new();
    authOptions.AutoSignIn = (context, cancellationToken) => {
        if (context.Activity.Value == null)
        {
            return Task.FromResult(true);
        }
        string command = (context.Activity.Value as JObject).Value<string>("commandId");
        bool signOutActivity = command == "signOutCommand";
        if (signOutActivity)
        {
            return Task.FromResult(false);
        }
        return Task.FromResult(true);
    };
    authOptions.AddAuthentication("graph", new TeamsSsoSettings(new string[] { "Files.Read.All" }, signInLink, msal));

    Application<AppState> app = new ApplicationBuilder<AppState>()
        .WithAIOptions(options)
        .WithStorage(sp.GetService<IStorage>())
        .WithTurnStateFactory(() => new AppState())
        .WithAuthentication(adapter, authOptions)
        .Build();

    app.OnConversationUpdate("membersAdded", async (turnContext, turnState, cancellationToken) =>
    {
        var welcomeText = "How can I help you today?";
        foreach (var member in turnContext.Activity.MembersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText), cancellationToken);
            }
        }
    });
    
    app.Authentication.Get("graph").OnUserSignInSuccess(async (turnContext, turnState) =>
    {
        // Successfully logged in
        await turnContext.SendActivityAsync("You are successfully logged in. You can send a new message to talk to the bot.");
    });
    app.Authentication.Get("graph").OnUserSignInFailure(async (turnContext, turnState, error) =>
    {
        // Failed to login
        await turnContext.SendActivityAsync("Failed to login");
        await turnContext.SendActivityAsync($"Error message: { error.Message}");
    });
    app.OnFeedbackLoop((turnContext, turnState, feedbackLoopData, _) =>
    {
        Console.WriteLine($"Your feedback is {turnContext.Activity.Value.ToString()}");
        return Task.CompletedTask;
    });
    
    return app;
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseStaticFiles();
app.UseRouting();
app.MapControllers();

app.Run();