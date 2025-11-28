using System.Text.Json;
using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.Summarizer
{
    /// <summary>
    /// Summarizer prompt configuration and functions
    /// </summary>
    public static class SummarizerPrompt
    {
        /// <summary>
        /// Get the summarizer prompt instructions
        /// </summary>
        public static string GetPromptInstructions()
        {
            return @"You are the Summarizer capability of the Collaborator that specializes in analyzing conversations between groups of people.
Your job is to retrieve and analyze conversation messages, then provide structured summaries with proper attribution.

<TIMEZONE AWARENESS>
The system uses the user's actual timezone from Microsoft Teams for all time calculations.
Time ranges will be pre-calculated by the Manager and passed to you as ISO timestamps when needed.

<INSTRUCTIONS>
1. Use the appropriate function to retrieve the messages you need based on the user's request
2. If time ranges are specified in the request, they will be pre-calculated and provided as ISO timestamps
3. If no specific timespan is mentioned, default to the last 24 hours using get_messages_by_time_range
4. Analyze the retrieved messages and identify participants and topics
5. Return a BRIEF summary with proper participant attribution
6. Include participant names in your analysis and summary points
7. Be concise and focus on the key topics discussed

<OUTPUT FORMAT>
- Use bullet points for main topics
- Include participant names when attributing ideas or statements
- Provide a brief overview if requested";
        }

        /// <summary>
        /// Register summarizer functions to the prompt
        /// </summary>
        public static void RegisterFunctions(
            OpenAIChatPrompt prompt,
            MessageContext context,
            ILogger logger
        )
        {
            // Register summarize_conversation function
            prompt.Function(
                "summarize_conversation",
                "Summarize the conversation history from the database within the specified time range",
                async () =>
                {
                    logger.LogDebug(
                        "?? Summarizer Capability - Start Time: {StartTime}",
                        context.StartTime
                    );
                    logger.LogDebug(
                        "?? Summarizer Capability - End Time: {EndTime}",
                        context.EndTime
                    );

                    // Get messages from database by time range
                    var allMessages = await context.Memory.GetMessagesByTimeRangeAsync(
                        context.StartTime,
                        context.EndTime
                    );

                    logger.LogDebug(
                        "?? Retrieved {Count} messages from database",
                        allMessages.Count
                    );

                    // Format conversation data
                    var conversationData = new
                    {
                        members = context.Members.Select(m => m.Name).ToList(),
                        timeRange = new { start = context.StartTime, end = context.EndTime },
                        messageCount = allMessages.Count,
                        messages = allMessages
                            .Select(m => new
                            {
                                role = m.Role,
                                name = m.Name,
                                content = m.Content,
                                timestamp = m.Timestamp,
                            })
                            .ToList(),
                    };

                    return JsonSerializer.Serialize(
                        conversationData,
                        new JsonSerializerOptions { WriteIndented = true }
                    );
                }
            );
        }
    }
}
