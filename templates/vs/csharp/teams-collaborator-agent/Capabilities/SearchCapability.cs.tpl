using {{SafeProjectName}}.Agent;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Text.Json;

namespace {{SafeProjectName}}.Capabilities;

public class SearchCapability : ICapability
{
    public string Name => "search";
    
    public string Description => @"**Search**: Use for requests like:
- 'find', 'search', 'look for', 'when did', 'who said'
- 'locate message', 'recall discussion'";

    public async Task<string> ExecuteAsync(AgentContext context, Kernel kernel)
    {
        var chatCompletion = kernel.GetRequiredService<IChatCompletionService>();
        
        // Get all conversation messages (search entire history)
        var messages = await context.Storage.GetAllMessagesAsync(context.ConversationId);

        if (!messages.Any())
        {
            return "No messages found in this conversation.";
        }

        // Build context for semantic search
        var messagesJson = JsonSerializer.Serialize(messages.Select(m => new
        {
            timestamp = m.Timestamp,
            name = m.Name,
            content = m.Content
        }));

        var prompt = $@"You are a conversation search assistant. The user is trying to find specific information from past conversations.

**User Query:** {context.UserMessage}

**Conversation History:**
```json
{messagesJson}
```

**Instructions:**
1. Search through the conversation history for relevant information
2. Find messages that match the user's search intent
3. Present the most relevant findings with context
4. Include timestamps and who said what
5. If no relevant messages found, say so clearly

Provide search results in a clear, easy-to-read format with quotes from the actual messages.";

        var chatHistory = new ChatHistory();
        chatHistory.AddSystemMessage(prompt);
        chatHistory.AddUserMessage("Find the information requested above.");

        var result = await chatCompletion.GetChatMessageContentAsync(chatHistory);
        return result.Content ?? "Unable to perform search.";
    }
}
