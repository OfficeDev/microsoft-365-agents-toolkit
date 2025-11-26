using System.Text.Json;
using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.Template
{
    /// <summary>
    /// Template prompt configuration and functions
    /// 
    /// TEMPLATE INSTRUCTIONS:
    /// This file defines the system prompt and registers functions for your capability.
    /// Follow these steps to customize it:
    /// 
    /// 1. Update GetPromptInstructions() with your capability's specific instructions
    /// 2. In RegisterFunctions(), register your custom functions using prompt.Function()
    /// 3. Each function can access context.Memory to retrieve/store conversation data
    /// </summary>
    public static class TemplatePrompt
    {
        /// <summary>
        /// Get the template prompt instructions
        /// 
        /// TODO: Replace this with your capability's system prompt
        /// This tells the AI model what role it plays and how to behave
        /// </summary>
        public static string GetPromptInstructions()
        {
            return @"You are the [CAPABILITY_NAME] capability of the Collaborator bot.

<ROLE>
TODO: Describe what this capability does and when it should be used

<INSTRUCTIONS>
1. TODO: List the steps this capability should follow
2. TODO: Explain how it should process user requests
3. TODO: Define the expected output format

<OUTPUT FORMAT>
- TODO: Specify how results should be formatted
- TODO: Include examples if helpful

<NOTES>
- TODO: Any additional guidelines or constraints
- Be clear, helpful, and concise.";
        }

        /// <summary>
        /// Register template functions to the prompt
        /// 
        /// TODO: Register your capability's functions here
        /// Each function can be called by the AI model during conversation
        /// </summary>
        public static void RegisterFunctions(
            OpenAIChatPrompt prompt,
            MessageContext context,
            ILogger logger
        )
        {
            // EXAMPLE 1: Function with NO parameters (like summarizer/actionItems)
            // Uncomment and customize this if your function doesn't need parameters:
            /*
            prompt.Function(
                "my_function_name",
                "Description of what this function does",
                async () =>
                {
                    logger.LogDebug("🔧 Template - Executing my_function_name");
                    
                    // Access conversation messages from the database
                    var messages = await context.Memory.GetMessagesByTimeRangeAsync(
                        context.StartTime,
                        context.EndTime
                    );
                    
                    logger.LogDebug("🔧 Retrieved {Count} messages", messages.Count);
                    
                    // TODO: Process the messages and return data
                    var result = new
                    {
                        messageCount = messages.Count,
                        data = messages.Select(m => new { m.Name, m.Content })
                    };
                    
                    return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
                }
            );
            */

            // EXAMPLE 2: Function WITH parameters (like search)
            // Uncomment and customize this if your function needs parameters:
            /*
            prompt.Function(
                "my_function_with_params",
                "Description of what this function does",
                TemplateSchema.GetMyFunctionSchema(), // Reference the schema you defined
                async (string param1, string[] param2, int param3 = 10) =>
                {
                    logger.LogDebug("🔧 Template - param1: {Param1}", param1);
                    logger.LogDebug("🔧 Template - param2: {Param2}", string.Join(", ", param2));
                    
                    // TODO: Use the parameters to process data
                    var filteredMessages = await context.Memory.GetFilteredMessagesAsync(
                        param2,
                        context.StartTime,
                        context.EndTime,
                        maxResults: param3
                    );
                    
                    // TODO: Return formatted results
                    return $"Found {filteredMessages.Count} messages matching criteria";
                }
            );
            */

            // TODO: Add your function registration here
            logger.LogWarning("⚠️ Template capability has no functions registered yet. Please implement your functions.");
        }
    }
}
