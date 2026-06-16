using Azure.Core;
using Azure.Identity;
using {{SafeProjectName}};
using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Capability.ActionItems;
using {{SafeProjectName}}.Capability.Search;
using {{SafeProjectName}}.Capability.Summarizer;
using {{SafeProjectName}}.Controllers;
using {{SafeProjectName}}.Storage;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.Api.Auth;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Extensions;
using Microsoft.Teams.Extensions.Logging;
using Microsoft.Teams.Plugins.AspNetCore.Extensions;

using TeamsLogging = Microsoft.Teams.Common.Logging;

var builder = WebApplication.CreateBuilder(args);
var config = builder.Configuration.Get<ConfigOptions>();

if (config == null)
{
    throw new InvalidOperationException("Missing configuration for ConfigOptions");
}

var teamsLogger = new TeamsLogging.ConsoleLogger("collaborator", TeamsLogging.LogLevel.Debug);
builder.Logging.AddTeams(teamsLogger);
builder.Services.AddSingleton<TeamsLogging.ILogger>(teamsLogger);

ConfigHelper.Initialize(config);
var appBuilder = App.Builder(new AppOptions { Logger = teamsLogger });

if (config.Teams.BotType == "UserAssignedMsi")
{
    appBuilder.AddCredentials(
        new TokenCredentials(
            config.Teams.ClientId ?? string.Empty,
            async (tenantId, scopes) =>
            {
                var clientId = config.Teams.ClientId;
                var managedIdentityCredential = new ManagedIdentityCredential(clientId);
                var tokenRequestContext = new TokenRequestContext(scopes, tenantId: tenantId);
                var accessToken = await managedIdentityCredential.GetTokenAsync(tokenRequestContext);
                return new TokenResponse { TokenType = "Bearer", AccessToken = accessToken.Token };
            }
        )
    );
}

builder.Services.AddSingleton<IDatabase>(sp =>
{
    var loggerFactory = sp.GetRequiredService<ILoggerFactory>();
    var dbLogger = loggerFactory.CreateLogger("Database");
    return StorageFactory.CreateStorageAsync(dbLogger).GetAwaiter().GetResult();
});

builder.Services.AddSingleton<List<CapabilityDefinition>>(sp =>
{
    var loggerFactory = sp.GetRequiredService<ILoggerFactory>();

    var summarizerLogger = loggerFactory.CreateLogger<SummarizerCapability>();
    var actionItemsLogger = loggerFactory.CreateLogger<ActionItemsCapability>();
    var searchLogger = loggerFactory.CreateLogger<SearchCapability>();

    var capabilities = new List<CapabilityDefinition>
    {
        SummarizerCapability.CreateDefinition(summarizerLogger),
        ActionItemsCapability.CreateDefinition(actionItemsLogger),
        SearchCapability.CreateDefinition(searchLogger),

        // ============================================================================
        // TO ADD A NEW CAPABILITY:
        // ============================================================================
        // 1. Copy the Capability/Template folder and rename it (e.g., "MyCapability").
        // 2. Update the capability files following the TODO comments in Template.
        // 3. Add a using statement at the top of this file:
        //    using {{SafeProjectName}}.Capability.MyCapability;
        // 4. Create a logger for your capability:
        //    var myCapabilityLogger = loggerFactory.CreateLogger<MyCapability>();
        // 5. Add your capability definition to this list:
        //    MyCapability.CreateDefinition(myCapabilityLogger),
        // 6. Update appsettings.json to include model config for your capability.
        // ============================================================================

        // EXAMPLE - Uncomment these lines to enable the Template capability:
        // var templateLogger = loggerFactory.CreateLogger<TemplateCapability>();
        // TemplateCapability.CreateDefinition(templateLogger),
    };

    return capabilities;
});

builder.Services.AddSingleton<Controller>();

builder.AddTeams(appBuilder, skipAuth: string.IsNullOrEmpty(config.Teams.ClientId));
var app = builder.Build();
var logger = app.Services.GetRequiredService<ILogger<Program>>();
ConfigHelper.ValidateEnvironment(logger);
ConfigHelper.LogModelConfigs(logger);
app.UseTeams();
app.Run();
