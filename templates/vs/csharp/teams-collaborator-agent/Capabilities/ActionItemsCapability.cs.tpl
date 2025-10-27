using {{SafeProjectName}}.Agent;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Text.Json;

namespace {{SafeProjectName}}.Capabilities;

public class ActionItemsCapability : ICapability
{
    public string Name => "action_items";
    
    public string Description => @"**Action Items**: Use for requests like:
- 'next steps', 'to-do', 'assign task', 'my tasks', 'what needs to be done'";

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

        // Build list of participants
        var participants = string.Join(", ", context.Members.Select(m => m.Name));

        // Build context for action item extraction
        var messagesJson = JsonSerializer.Serialize(messages.Select(m => new
        {
            timestamp = m.Timestamp,
            name = m.Name,
            content = m.Content
        }));

        var prompt = $@"You are an action item tracker. Analyze the following conversation and extract all action items, tasks, and commitments.

**Participants:** {participants}

**Conversation Messages:**
```json
{messagesJson}
```

**Instructions:**
1. Extract all explicit action items and tasks mentioned
2. Identify who is responsible for each task (if mentioned)
3. Note any deadlines or timelines (if mentioned)
4. Include follow-up items and next steps
5. Format as a clear, actionable list

Provide the action items in this format:
- [ ] **[Person]**: Task description (Deadline: if mentioned)

If no action items found, respond with: "No action items identified in this conversation."";

        var chatHistory = new ChatHistory();
        chatHistory.AddSystemMessage(prompt);
        chatHistory.AddUserMessage(context.UserMessage);

        var result = await chatCompletion.GetChatMessageContentAsync(chatHistory);
        return result.Content ?? "Unable to extract action items.";
    }
}
