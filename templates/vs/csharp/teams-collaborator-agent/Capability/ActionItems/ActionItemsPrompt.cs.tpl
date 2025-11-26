using System.Text.Json;
using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.ActionItems
{
    /// <summary>
    /// Action Items prompt configuration and functions
    /// </summary>
    public static class ActionItemsPrompt
    {
        /// <summary>
        /// Get the action items prompt instructions
        /// </summary>
        public static string GetPromptInstructions()
        {
            return @"You are the Action Items capability of the Collaborator bot. Your role is to analyze team conversations and extract a list of clear action items based on what people said.

<GOAL>
Your job is to generate a concise, readable list of action items mentioned in the conversation. Focus on identifying:
- What needs to be done
- Who will do it (if mentioned)

<EXAMPLES OF ACTION ITEM CLUES>
- ""I'll take care of this""
- ""Can you follow up on...""
- ""Let's finish this by tomorrow""
- ""We still need to decide...""
- ""Assign this to Alex""
- ""We should check with finance""

<OUTPUT FORMAT>
- Return a plain text list of bullet points
- Each item should include a clear task and a person (if known)

<EXAMPLE OUTPUT>
- ✅ Sarah will create the draft proposal by Friday
- ✅ Alex will check budget numbers before the meeting
- ✅ Follow up with IT on access issues
- ✅ Decide final presenters by end of week

<NOTES>
- If no one is assigned, just describe the task
- Skip greetings or summary text — just the action items
- Do not assign tasks unless the conversation suggests it

Be clear, helpful, and concise.";
        }

        /// <summary>
        /// Register action items functions to the prompt
        /// </summary>
        public static void RegisterFunctions(
            OpenAIChatPrompt prompt,
            MessageContext context,
            ILogger logger
        )
        {
            // Register generate_action_items function
            prompt.Function(
                "generate_action_items",
                "Generate a list of action items based on the conversation",
                async () =>
                {
                    logger.LogDebug(
                        "📋 Action Items Capability - Start Time: {StartTime}",
                        context.StartTime
                    );
                    logger.LogDebug(
                        "📋 Action Items Capability - End Time: {EndTime}",
                        context.EndTime
                    );

                    // Get messages from database by time range
                    var allMessages = await context.Memory.GetMessagesByTimeRangeAsync(
                        context.StartTime,
                        context.EndTime
                    );

                    logger.LogDebug(
                        "📋 Retrieved {Count} messages for action item extraction",
                        allMessages.Count
                    );

                    // Format messages data
                    var messagesData = new
                    {
                        messages = allMessages
                            .Select(m => new
                            {
                                timestamp = m.Timestamp,
                                name = m.Name,
                                content = m.Content,
                            })
                            .ToList(),
                    };

                    return JsonSerializer.Serialize(
                        messagesData,
                        new JsonSerializerOptions { WriteIndented = true }
                    );
                }
            );
        }
    }
}
