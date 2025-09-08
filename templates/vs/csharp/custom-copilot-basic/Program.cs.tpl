using {{SafeProjectName}};
using {{SafeProjectName}}.Controllers;
using Azure.Core;
using Azure.Identity;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Models.OpenAI.Extensions;
using Microsoft.Teams.AI.Prompts;
using Microsoft.Teams.Api.Auth;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Extensions;
using Microsoft.Teams.Common.Http;
using Microsoft.Teams.Plugins.AspNetCore.Extensions;

var builder = WebApplication.CreateBuilder(args);
var config = builder.Configuration.Get<ConfigOptions>();

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
    Console.WriteLine($"Using UserAssignedMSI");
    appBuilder.AddCredentials(new TokenCredentials(
        config.Teams.ClientId ?? string.Empty,
        async (tenantId, scopes) =>
        {
            return await createTokenFactory(scopes, tenantId);
        }
    ));
}

builder.Services.AddSingleton<Controller>();
builder.AddTeams(appBuilder);

// Read instructions from file
var instructionsPath = Path.Combine(builder.Environment.ContentRootPath, "Prompts", "instructions.txt");
var instructions = await File.ReadAllTextAsync(instructionsPath);

builder.Services.AddOpenAI(
    {{#useOpenAI}}
    new OpenAIChatModel(
        config.OpenAI.ApiKey, 
        config.OpenAI.DefaultModel),
    {{/useOpenAI}}
    {{#useAzureOpenAI}}
    new OpenAIChatModel(
        config.Azure.OpenAIDeploymentName, 
        config.Azure.OpenAIApiKey,
        new() { Endpoint = new Uri($"{config.Azure.OpenAIEndpoint}/openai/v1") }),
    {{/useAzureOpenAI}}
    new ChatPromptOptions().WithInstructions(instructions));

var app = builder.Build();
var teamsApp = app.UseTeams();
teamsApp.OnError(async (sender, @event) =>
{
    var context = @event.Context;
    await context.Send($"The agent encountered an unhandled error: {@event.Exception.Message}");
    await context.Send($"To continue to run this agent, please fix the agent source code.");
    teamsApp.Logger.Error($"[OnError] unhandled error : {@event.Exception.Message}");
});

app.Run();