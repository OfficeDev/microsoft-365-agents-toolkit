using {{SafeProjectName}};
using {{SafeProjectName}}.Agent;
using {{SafeProjectName}}.Capabilities;
using {{SafeProjectName}}.Storage;
using Azure.Core;
using Azure.Identity;
using Microsoft.Teams.Api.Auth;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Extensions;
using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.SemanticKernel;

var builder = WebApplication.CreateBuilder(args);
var config = builder.Configuration.Get<ConfigOptions>();

if (config == null)
{
    throw new InvalidOperationException("Missing configuration for ConfigOptions");
}

// Configure Token Credentials for Managed Identity
Func<string[], string?, Task<ITokenResponse>> createTokenFactory = async (string[] scopes, string? tenantId) =>
{
    var clientId = config.Teams.ClientId;
    var managedIdentityCredential = new ManagedIdentityCredential(clientId);
    var tokenRequestContext = new TokenRequestContext(scopes, tenantId: tenantId);
    var accessToken = await managedIdentityCredential.GetTokenAsync(tokenRequestContext);

    return new TokenResponse
    {
        TokenType = "Bearer",
        AccessToken = accessToken.Token,
    };
};

var appBuilder = App.Builder();

if (config.Teams.BotType == "UserAssignedMsi")
{
    appBuilder.AddCredentials(new TokenCredentials(
        config.Teams.ClientId ?? string.Empty,
        async (tenantId, scopes) => await createTokenFactory(scopes, tenantId)
    ));
}

builder.AddTeams(appBuilder);

// Configure Semantic Kernel
var kernelBuilder = Kernel.CreateBuilder();

{{#useOpenAI}}
kernelBuilder.AddOpenAIChatCompletion(
    modelId: config.OpenAI.DefaultModel,
    apiKey: config.OpenAI.ApiKey);
{{/useOpenAI}}
{{#useAzureOpenAI}}
kernelBuilder.AddAzureOpenAIChatCompletion(
    deploymentName: config.Azure.OpenAIDeploymentName,
    endpoint: config.Azure.OpenAIEndpoint,
    apiKey: config.Azure.OpenAIApiKey);
{{/useAzureOpenAI}}

var kernel = kernelBuilder.Build();
builder.Services.AddSingleton(kernel);

// Configure Storage
builder.Services.AddSingleton<IConversationStorage>(sp =>
{
    var storageType = config.Storage.Type.ToLowerInvariant();
    return storageType switch
    {
        "sqlite" => new SqliteConversationStorage(config.Storage.ConnectionString),
        "sqlserver" => new SqlServerConversationStorage(config.Storage.ConnectionString),
        _ => throw new InvalidOperationException($"Unsupported storage type: {storageType}")
    };
});

// Register Capabilities
builder.Services.AddSingleton<CapabilityRegistry>();
builder.Services.AddSingleton<ICapability, SummarizerCapability>();
builder.Services.AddSingleton<ICapability, ActionItemsCapability>();
builder.Services.AddSingleton<ICapability, SearchCapability>();

// Register Agent Manager
builder.Services.AddSingleton<AgentManager>();

var app = builder.Build();
app.UseTeams();
app.Run();
