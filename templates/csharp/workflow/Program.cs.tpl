using {{SafeProjectName}};
using {{SafeProjectName}}.Commands;
using {{SafeProjectName}}.CardActions;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using Microsoft.TeamsFx.Conversation;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();
builder.Services.AddCloudAdapter();
builder.Logging.AddConsole();

// Add AspNet token validation
builder.Services.AddBotAspNetAuthentication(builder.Configuration);

// Register IStorage.  For development, MemoryStorage is suitable.
// For production Agents, persisted storage should be used so
// that state survives Agent restarts, and operate correctly
// in a cluster of Agent instances.
builder.Services.AddSingleton<IStorage, MemoryStorage>();

// Add AgentApplicationOptions from config.
builder.AddAgentApplicationOptions();

// Create command handlers
builder.Services.AddSingleton<HelloWorldCommandHandler>();
builder.Services.AddSingleton<GenericCommandHandler>();
builder.Services.AddSingleton<DoStuffActionHandler>();

// Keep the ConversationBot to maintain compatibility with TeamsFx SDK
builder.Services.AddSingleton(sp =>
{
    var options = new ConversationOptions()
    {
        Adapter = sp.GetService<CloudAdapter>(),
        Command = new CommandOptions()
        {
            Commands = new List<ITeamsCommandHandler> { 
                sp.GetService<HelloWorldCommandHandler>(), 
                sp.GetService<GenericCommandHandler>() 
            }
        },
        CardAction = new CardActionOptions()
        {
            Actions = new List<IAdaptiveCardActionHandler> { sp.GetService<DoStuffActionHandler>() }
        }
    };
    return new ConversationBot(options);
});

// Add the bot (which is transient)
builder.AddAgent<TeamsBot>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

// Map the /api/messages endpoint to the AgentApplication
app.MapPost("/api/messages", async (HttpRequest request, HttpResponse response, IAgentHttpAdapter adapter, IAgent agent, CancellationToken cancellationToken) =>
{
    await adapter.ProcessAsync(request, response, agent, cancellationToken);
});

if (app.Environment.IsDevelopment() || app.Environment.EnvironmentName == "TestTool")
{
    app.MapGet("/", () => "Workflow Bot");
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();