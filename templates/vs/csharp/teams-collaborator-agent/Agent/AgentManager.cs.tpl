using {{SafeProjectName}}.Capabilities;
using {{SafeProjectName}}.Storage;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Text;

namespace {{SafeProjectName}}.Agent;

public class AgentManager
{
    private readonly Kernel _kernel;
    private readonly ICapabilityRegistry _capabilityRegistry;
    private readonly IConversationStorage _storage;
    private readonly ILogger<AgentManager> _logger;

    public AgentManager(
        Kernel kernel,
        ICapabilityRegistry capabilityRegistry,
        IConversationStorage storage,
        ILogger<AgentManager> logger)
    {
        _kernel = kernel;
        _capabilityRegistry = capabilityRegistry;
        _storage = storage;
        _logger = logger;
    }

    public async Task<string> ProcessMessageAsync(AgentContext context)
    {
        try
        {
            _logger.LogInformation($"Processing message for conversation {context.ConversationId}");

            // Determine which capability to use based on the user's message
            var capability = await DetermineCapabilityAsync(context);

            if (capability == null)
            {
                return "I'm not sure how to help with that. You can ask me to:\n" +
                       "- Summarize the conversation\n" +
                       "- Extract action items\n" +
                       "- Search for specific information";
            }

            _logger.LogInformation($"Selected capability: {capability.Name}");

            // Execute the capability
            var result = await capability.ExecuteAsync(context, _kernel);
            
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing message");
            return "I encountered an error while processing your request. Please try again.";
        }
    }

    private async Task<ICapability?> DetermineCapabilityAsync(AgentContext context)
    {
        var chatCompletion = _kernel.GetRequiredService<IChatCompletionService>();
        
        // Build capability descriptions
        var capabilitiesDesc = new StringBuilder();
        foreach (var cap in _capabilityRegistry.GetAllCapabilities())
        {
            capabilitiesDesc.AppendLine(cap.Description);
        }

        var prompt = $@"You are a routing agent. Analyze the user's message and determine which capability should handle it.

**Available Capabilities:**
{capabilitiesDesc}

**User Message:** {context.UserMessage}

**Instructions:**
Respond with ONLY the capability name that best matches the user's intent:
- 'summarizer' for summary requests
- 'action_items' for action item extraction
- 'search' for searching conversation history
- 'none' if the request doesn't match any capability

Response (one word only):";

        var chatHistory = new ChatHistory();
        chatHistory.AddSystemMessage(prompt);
        chatHistory.AddUserMessage("Which capability?");

        var result = await chatCompletion.GetChatMessageContentAsync(chatHistory);
        var capabilityName = result.Content?.Trim().ToLowerInvariant();

        _logger.LogDebug($"Determined capability: {capabilityName}");

        if (string.IsNullOrEmpty(capabilityName) || capabilityName == "none")
        {
            return null;
        }

        return _capabilityRegistry.GetCapability(capabilityName);
    }

    public async Task<bool> RecordMessageAsync(
        string conversationId,
        string messageId,
        string name,
        string content,
        string role = "user")
    {
        try
        {
            var message = new Storage.Models.ConversationMessage
            {
                Id = messageId,
                ConversationId = conversationId,
                Name = name,
                Content = content,
                Timestamp = DateTime.UtcNow,
                Role = role
            };

            return await _storage.AddMessageAsync(message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error recording message");
            return false;
        }
    }

    public async Task<bool> ClearConversationHistoryAsync(string conversationId)
    {
        try
        {
            return await _storage.ClearConversationAsync(conversationId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error clearing conversation history");
            return false;
        }
    }
}
