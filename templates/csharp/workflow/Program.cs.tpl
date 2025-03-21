using {{SafeProjectName}};
using {{SafeProjectName}}.Commands;
using {{SafeProjectName}}.CardActions;
using Microsoft.Agents.BotBuilder.App;
using Microsoft.Agents.BotBuilder.State;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using Microsoft.TeamsFx.Conversation;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();
builder.Services.AddCloudAdapter<AdapterWithErrorHandler>();
builder.Logging.AddConsole();

// Add AspNet token validation
builder.Services.AddBotAspNetAuthentication(builder.Configuration);

// Add ApplicationOptions
builder.Services.AddTransient(sp =>
{
    return new AgentApplicationOptions()
    {
        StartTypingTimer = false,
        TurnStateFactory = () => new TurnState(sp.GetService<IStorage>())
    };
});

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
builder.AddBot<TeamsBot>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

if (app.Environment.IsDevelopment())
{
    app.MapGet("/", () => "Workflow Bot");
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();