using {{SafeProjectName}}.Agent;
using {{SafeProjectName}}.Storage;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Text;
using System.Text.Json;

namespace {{SafeProjectName}}.Capabilities;

public class SummarizerCapability : ICapability
{
    public string Name => "summarizer";
    
    public string Description => @"**Summarizer**: Use for keywords like:
- 'summarize', 'overview', 'recap', 'conversation history'
- 'what did we discuss', 'catch me up', 'who said what', 'recent messages'";

    public async Task<string> ExecuteAsync(AgentContext context, Kernel kernel)
    {
        var chatCompletion = kernel.GetRequiredService<IChatCompletionService>();
        
        // Get conversation messages from storage
        var messages = await context.Storage.GetMessagesByTimeRangeAsync(
            context.ConversationId,
            context.StartTime,
            context.EndTime
        );

        if (!messages.Any())
        {
            return "No messages found in the specified time range.";
        }

        // Build context for summarization
        var messagesJson = JsonSerializer.Serialize(messages.Select(m => new
        {
            timestamp = m.Timestamp,
            name = m.Name,
            content = m.Content
        }));

        var prompt = $@"You are a conversation summarizer. Analyze the following conversation and provide a clear, concise summary.

**Conversation Messages:**
```json
{messagesJson}
```

**Instructions:**
1. Identify the main topics discussed
2. Note key decisions or agreements
3. Highlight any outstanding questions or concerns
4. Keep the summary concise but informative

Provide the summary in a friendly, professional tone.";

        var chatHistory = new ChatHistory();
        chatHistory.AddSystemMessage(prompt);
        chatHistory.AddUserMessage(context.UserMessage);

        var result = await chatCompletion.GetChatMessageContentAsync(chatHistory);
        return result.Content ?? "Unable to generate summary.";
    }
}
