using {{SafeProjectName}};
using {{SafeProjectName}}.Bot.Agents;
using Microsoft.SemanticKernel;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Storage;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();
builder.Logging.AddConsole();


// Register Semantic Kernel
builder.Services.AddKernel();

// Register the AI service of your choice. AzureOpenAI and OpenAI are demonstrated...
var config = builder.Configuration.Get<ConfigOptions>();

{{#useAzureOpenAI}}
builder.Services.AddAzureOpenAIChatCompletion(
    deploymentName: config.Azure.OpenAIDeploymentName,
    endpoint: config.Azure.OpenAIEndpoint,
    apiKey: config.Azure.OpenAIApiKey
);
{{/useAzureOpenAI}}
{{#useOpenAI}}
builder.Services.AddOpenAIChatCompletion(
   modelId: config.OpenAI.DefaultModel,
   apiKey: config.OpenAI.ApiKey
);
{{/useOpenAI}}

// Register the WeatherForecastAgent
builder.Services.AddTransient<WeatherForecastAgent>();

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
builder.AddAgent<{{SafeProjectName}}.Bot.WeatherAgentBot>();

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

if (app.Environment.IsDevelopment() || app.Environment.EnvironmentName == "TestTool")
{
    app.MapGet("/", () => "Weather Bot");
    app.UseDeveloperExceptionPage();
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();

