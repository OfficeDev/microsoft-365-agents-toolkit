using {{SafeProjectName}};
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Teams.AI;
using Microsoft.Teams.AI.AI.Planners;
using Microsoft.Teams.AI.AI.Planners.Experimental;
using Microsoft.Teams.AI.AI;

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
{{#useOpenAI}}
builder.Services.AddSingleton(_ => new AssistantsPlannerOptions(config.OpenAI.ApiKey, config.OpenAI.AssistantId));
{{/useOpenAI}}
{{#useAzureOpenAI}}
builder.Services.AddSingleton(_ => new AssistantsPlannerOptions(config.Azure.OpenAIApiKey, config.Azure.OpenAIAssistantId, config.Azure.OpenAIEndpoint));
{{/useAzureOpenAI}}

// Create the bot as transient. In this case the ASP Controller is expecting an IBot.
builder.Services.AddTransient<IBot>(sp =>
{
    // Create loggers
    ILoggerFactory loggerFactory = sp.GetService<ILoggerFactory>();

    IPlanner<AssistantsState> planner = new AssistantsPlanner<AssistantsState>(sp.GetService<AssistantsPlannerOptions>()!, loggerFactory);
    AIOptions<AssistantsState> options = new(planner);
    options.EnableFeedbackLoop = true;

    Application<AssistantsState> app = new ApplicationBuilder<AssistantsState>()
        .WithAIOptions(options)
        .WithStorage(sp.GetService<IStorage>())
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

    app.AI.ImportActions(new ActionHandlers());
    app.OnFeedbackLoop((turnContext, turnState, feedbackLoopData, _) =>
    {
        Console.WriteLine($"Your feedback is {turnContext.Activity.Value.ToString()}");
        return Task.CompletedTask;
    });

    // Listen for user to say "/reset".
    app.OnMessage("/reset", ActivityHandlers.ResetMessageHandler);
    return app;
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseStaticFiles();
app.UseRouting();
app.UseEndpoints(endpoints =>
{
    endpoints.MapControllers();
});

app.Run();