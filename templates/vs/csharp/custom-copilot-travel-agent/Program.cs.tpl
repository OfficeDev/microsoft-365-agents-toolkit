{{#useAzureOpenAI}}
using Azure.AI.OpenAI;
using Azure;
{{/useAzureOpenAI}}
{{#useOpenAI}}
using OpenAI;
{{/useOpenAI}}
using {{SafeProjectName}};
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Storage;
using Microsoft.Extensions.AI;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();
builder.Logging.AddConsole();

var config = builder.Configuration.Get<ConfigOptions>();

builder.Services.AddSingleton(serviceProvider =>
{
{{#useAzureOpenAI}}
    return new AzureOpenAIClient(
        new Uri(config.Azure.OpenAIEndpoint),
        new AzureKeyCredential(config.Azure.OpenAIApiKey))
        .GetChatClient(config.Azure.OpenAIDeploymentName)
        .AsIChatClient();
{{/useAzureOpenAI}}
{{#useOpenAI}}
    return new OpenAIClient(config.OpenAI.ApiKey)
        .GetChatClient(config.OpenAI.DefaultModel)
        .AsIChatClient();
{{/useOpenAI}}
});

// Register the TravelAgentBot
builder.Services.AddTransient<{{SafeProjectName}}.Bot.TravelAgentBot>();

// Add AspNet token validation
builder.Services.AddBotAspNetAuthentication(builder.Configuration);

// Register IStorage.  For development, MemoryStorage is suitable.
// For production Agents, persisted storage should be used so
// that state survives Agent restarts, and operate correctly
// in a cluster of Agent instances.
builder.Services.AddSingleton<IStorage, MemoryStorage>();

// Add AgentApplicationOptions from config.
builder.AddAgentApplicationOptions();

// Add AgentApplicationOptions.  This will use DI'd services and IConfiguration for construction.
builder.Services.AddTransient<AgentApplicationOptions>();

// Add the bot (which is transient)
builder.AddAgent<{{SafeProjectName}}.Bot.TravelAgentBot>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapPost("/api/messages", async (HttpRequest request, HttpResponse response, IAgentHttpAdapter adapter, IAgent agent, CancellationToken cancellationToken) =>
{
    await adapter.ProcessAsync(request, response, agent, cancellationToken);
});

if (app.Environment.IsDevelopment())
{
    app.MapGet("/", () => "{{ProjectName}}");
    app.UseDeveloperExceptionPage();
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();

