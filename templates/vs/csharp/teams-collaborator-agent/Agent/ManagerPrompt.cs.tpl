using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Agent
{
    /// <summary>
    /// Manager prompt configuration and functions
    /// </summary>
    public static class ManagerPrompt
    {
        /// <summary>
        /// Get the manager prompt template with placeholders
        /// </summary>
        public static string GetPromptTemplate()
        {
            return @"You are the Manager for the Collaborator — a Microsoft Teams bot. You coordinate requests by deciding which specialized capability should handle each @mention.

<AVAILABLE CAPABILITIES>
{capabilityList}

<INSTRUCTIONS>
1. **If the request includes a time expression**, call calculate_time_range first using the exact phrase (e.g., ""last week"", ""past 2 days"", ""today"").
2. Analyze the request’s intent and route it to the best-matching capability.
3. If no capability applies, respond conversationally and describe what Collaborator *can* help with.

<WHEN TO USE EACH CAPABILITY>
Use the following descriptions to determine routing logic. Match based on intent, not just keywords.

{capabilityDescriptions}

<RESPONSE RULE>
When using a function call to delegate, return the capability’s response **as-is**, with no added commentary or explanation. MAKE SURE TO NOT WRAP THE RESPONSE IN QUOTES.

✅ GOOD: [capability response]  
❌ BAD: Here’s what the Summarizer found: [capability response]

<GENERAL RESPONSES>
Be warm and helpful when the request is casual or unclear. Mention your abilities naturally.

✅ Hi there! I can help with summaries, task tracking, or finding specific messages.
✅ Interesting! I specialize in conversation analysis and action items. Want help with that?";
        }

        /// <summary>
        /// Register all manager functions to the prompt
        /// </summary>
        public static void RegisterFunctions(
            OpenAIChatPrompt prompt,
            MessageContext context,
            Manager manager,
            ILogger logger
        )
        {
            // Register calculate_time_range function
            prompt.Function(
                "calculate_time_range",
                "Parse natural language time expressions and calculate exact start/end times for time-based queries",
                (string time_phrase) =>
                {
                    logger.LogDebug(
                        $"?? FUNCTION CALL: calculate_time_range - parsing \"{time_phrase}\""
                    );
                    manager.CalculateTimeRange(context, time_phrase);
                    return $"Time range calculated: from {context.StartTime} to {context.EndTime}";
                }
            );

            // Register clear_conversation_history function
            prompt.Function(
                "clear_conversation_history",
                "Clear the conversation history in the database for the current conversation",
                async () =>
                {
                    logger.LogDebug("??? FUNCTION CALL: clear_conversation_history");
                    await manager.ClearConversationHistoryAsync(context);
                    return "Conversation history has been cleared successfully.";
                }
            );
        }

        /// <summary>
        /// Register capability delegation functions
        /// </summary>
        public static void RegisterCapabilityFunctions(
            OpenAIChatPrompt prompt,
            MessageContext context,
            Manager manager,
            ILogger logger
        )
        {
            foreach (var capability in manager.GetCapabilities())
            {
                var capabilityName = capability.Name;
                var capabilityDescription = capability.ManagerDescription;

                prompt.Function(
                    $"delegate_to_{capabilityName}",
                    $"Delegate to {capabilityName} capability: {capabilityDescription}",
                    async () =>
                    {
                        logger.LogInformation($"?? FUNCTION CALL: delegate_to_{capabilityName}");
                        var result = await manager.DelegateToCapability(capabilityName, context);
                        logger.LogDebug(
                            $"? Capability {capabilityName} completed, response length: {result?.Length ?? 0} chars"
                        );
                        return result;
                    }
                );
            }
        }
    }
}
