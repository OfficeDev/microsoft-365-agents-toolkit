using {{SafeProjectName}}.Agent;
using {{SafeProjectName}}.Storage;
using {{SafeProjectName}}.Storage.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Api.Models;
using System.Text.Json;

namespace {{SafeProjectName}}.Controllers;

[ApiController]
[Route("api/messages")]
public class BotController : ControllerBase
{
    private readonly AgentManager _agentManager;
    private readonly IConversationStorage _storage;
    private readonly ILogger<BotController> _logger;
    private readonly ConfigOptions _config;

    public BotController(
        AgentManager agentManager,
        IConversationStorage storage,
        ILogger<BotController> logger,
        ConfigOptions config)
    {
        _agentManager = agentManager;
        _storage = storage;
        _logger = logger;
        _config = config;
    }

    [HttpPost]
    public async Task<IActionResult> PostAsync([FromBody] Activity activity)
    {
        try
        {
            _logger.LogInformation($"Received activity: {activity.Type}");

            if (activity.Type == "message")
            {
                await HandleMessageActivityAsync(activity);
            }
            else if (activity.Type == "conversationUpdate" && activity.MembersAdded?.Any() == true)
            {
                await HandleInstallActivityAsync(activity);
            }

            return Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing activity");
            return StatusCode(500, "Internal server error");
        }
    }

    private async Task HandleMessageActivityAsync(Activity activity)
    {
        var conversationId = activity.Conversation?.Id ?? string.Empty;
        var userMessage = activity.Text ?? string.Empty;
        var from = activity.From?.Name ?? "Unknown";

        // Record the incoming user message
        await _agentManager.RecordMessageAsync(
            conversationId,
            activity.Id ?? Guid.NewGuid().ToString(),
            from,
            userMessage,
            "user");

        // Check if bot was mentioned (for group chats)
        var isBotMentioned = activity.Entities?.Any(e => e.Type == "mention") == true;
        var isGroupChat = activity.Conversation?.IsGroup == true;

        // Process if it's a 1:1 chat or bot was mentioned in group chat
        if (!isGroupChat || isBotMentioned)
        {
            // Build context
            var context = new AgentContext
            {
                ConversationId = conversationId,
                UserMessage = userMessage,
                StartTime = DateTime.UtcNow.AddHours(-24), // Default: last 24 hours
                EndTime = DateTime.UtcNow,
                Members = new List<ConversationMember>
                {
                    new() { Id = activity.From?.Id ?? "", Name = from, Role = "user" }
                },
                Storage = _storage,
                BotId = activity.Recipient?.Id ?? ""
            };

            // Process the message through the agent manager
            var response = await _agentManager.ProcessMessageAsync(context);

            // Send response back (this would normally go through Bot Framework adapter)
            _logger.LogInformation($"Bot response: {response}");

            // Record bot's response
            await _agentManager.RecordMessageAsync(
                conversationId,
                Guid.NewGuid().ToString(),
                "Collab Agent",
                response,
                "assistant");

            // TODO: Actually send the response via Bot Framework adapter
            // This would typically involve using the Bot Framework SDK's TurnContext
        }
    }

    private Task HandleInstallActivityAsync(Activity activity)
    {
        _logger.LogInformation("Bot was installed to conversation");
        
        // TODO: Send welcome message via Bot Framework adapter
        var welcomeMessage = "👋 Hi! I'm the Collab Agent 🚀. I'll listen to the conversation and can provide summaries, action items, or search for a message when asked!";
        
        _logger.LogInformation($"Welcome message: {welcomeMessage}");
        
        return Task.CompletedTask;
    }
}
