using System.Text.Json;
using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.Search
{
    /// <summary>
    /// Search prompt configuration and functions
    /// </summary>
    public static class SearchPrompt
    {
        /// <summary>
        /// Get the search prompt instructions
        /// </summary>
        public static string GetPromptInstructions()
        {
            return @"You are the Search capability of the Collaborator. Your role is to help users find specific conversations or messages from their chat history.

You can search through message history to find:
- Conversations between specific people
- Messages about specific topics
- Messages from specific time periods (time ranges will be pre-calculated by the Manager)
- Messages containing specific keywords

When a user asks you to find something, use the search_messages function to search the database.

RESPONSE FORMAT:
- Your search_messages function returns just the text associated with the search results
- Focus on creating a helpful, conversational summary that complements the citations
- Be specific about what was found and provide context about timing and participants
- If no results are found, suggest alternative search terms or broader criteria

Be helpful and conversational in your responses. The user will see both your text response and interactive cards that let them jump to the original messages.";
        }

        /// <summary>
        /// Register search functions to the prompt
        /// </summary>
        public static void RegisterFunctions(
            OpenAIChatPrompt prompt,
            MessageContext context,
            ILogger logger
        )
        {
            // Register search_messages function
            prompt.Function(
                "search_messages",
                "Search the conversation for relevant messages",
                async (
                    string[] keywords,
                    string[]? participants = null,
                    int max_results = 5
                ) =>
                {
                    logger.LogDebug(
                        "🔍 Search Capability - Keywords: {Keywords}",
                        string.Join(", ", keywords)
                    );

                    if (participants != null && participants.Length > 0)
                    {
                        logger.LogDebug(
                            "🔍 Search Capability - Participants filter: {Participants}",
                            string.Join(", ", participants)
                        );
                    }

                    // Get filtered messages from database
                    var selected = await context.Memory.GetFilteredMessagesAsync(
                        keywords,
                        context.StartTime,
                        context.EndTime,
                        participants,
                        max_results
                    );

                    logger.LogDebug("🔍 Found {Count} matching messages", selected.Count);

                    if (selected.Count == 0)
                    {
                        return "No matching messages found.";
                    }

                    // Create citations and format results
                    var results = new List<string>();

                    foreach (var msg in selected)
                    {
                        // Parse timestamp
                        DateTime parsedDate = DateTime.TryParse(msg.Timestamp, out var dt)
                            ? dt
                            : DateTime.UtcNow;
                        var date = parsedDate.ToString("g"); // Short date/time format

                        var preview =
                            msg.Content.Length > 100
                                ? msg.Content.Substring(0, 100) + "..."
                                : msg.Content;

                        // Create deep link to the message
                        var deepLink = CreateDeepLink(
                            msg.ActivityId ?? msg.Id,
                            context.ConversationId
                        );

                        // Create citation
                        var citation = new CitationAppearance
                        {
                            Title = $"Message from {msg.Name}",
                            Url = deepLink,
                            Content = $"{date}: \"{preview}\"",
                            AppearanceIndex = context.Citations.Count + 1,
                        };

                        context.Citations.Add(citation);

                        // Format result line
                        results.Add($"• [{msg.Name}]({deepLink}) at {date}: \"{preview}\"");
                    }

                    return string.Join("\n", results);
                }
            );
        }

        /// <summary>
        /// Create a Teams deep link to a specific message
        /// </summary>
        private static string CreateDeepLink(string activityId, string conversationId)
        {
            var contextParam = Uri.EscapeDataString(
                JsonSerializer.Serialize(new { contextType = "chat" })
            );
            var encodedConvId = Uri.EscapeDataString(conversationId);
            return $"https://teams.microsoft.com/l/message/{encodedConvId}/{activityId}?context={contextParam}";
        }
    }
}
