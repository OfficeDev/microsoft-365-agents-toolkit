using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.Search
{
    /// <summary>
    /// Search capability for finding specific messages in conversation history
    /// </summary>
    public class SearchCapability : BaseCapability
    {
        public override string Name => "search";

        public SearchCapability(ILogger logger)
            : base(logger) { }

        public override OpenAIChatPrompt CreatePrompt(MessageContext context)
        {
            var searchConfig = ConfigHelper.GetModelConfig("search");
            OpenAIChatModel model;
{{#useAzureOpenAI}}
            model = new OpenAIChatModel(
                searchConfig.Model,
                searchConfig.ApiKey,
                new() { Endpoint = new Uri($"{searchConfig.Endpoint}/openai/v1") }
            );
{{/useAzureOpenAI}}
{{#useOpenAI}}
            model = new OpenAIChatModel(
                searchConfig.ApiKey,
                searchConfig.Model
            );
{{/useOpenAI}}

            // Get instructions from SearchPrompt
            var instructions = SearchPrompt.GetPromptInstructions();
            var prompt = new OpenAIChatPrompt(
                model,
                new ChatPromptOptions().WithInstructions(instructions)
            );

            // Register functions using SearchPrompt
            SearchPrompt.RegisterFunctions(prompt, context, _logger);

            _logger.LogDebug("✅ Initialized Search Capability!");
            return prompt;
        }

        /// <summary>
        /// Create capability definition for manager registration
        /// </summary>
        public static CapabilityDefinition CreateDefinition(ILogger<SearchCapability> logger)
        {
            return new CapabilityDefinition
            {
                Name = "search",
                ManagerDescription =
                    @"**Search**: Use for:
- ""find"", ""search"", ""show me"", ""conversation with"", ""where did [person] say"", ""messages from last week""",
                Handler = async (context, handlerLogger) =>
                {
                    // Create Search capability (it will create its own ChatModel)
                    var capability = new SearchCapability(logger);
                    var result = await capability.ProcessRequestAsync(context);

                    if (!string.IsNullOrEmpty(result.Error))
                    {
                        handlerLogger.LogError("❌ Error in Search Capability: {Error}", result.Error);
                        return $"Error in Search Capability: {result.Error}";
                    }

                    return result.Response ?? "No response from Search Capability";
                },
            };
        }
    }
}
