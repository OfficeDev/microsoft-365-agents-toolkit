using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.Summarizer
{
    /// <summary>
    /// Summarizer capability for analyzing and summarizing conversations
    /// </summary>
    public class SummarizerCapability : BaseCapability
    {
        public override string Name => "summarizer";

        public SummarizerCapability(ILogger logger)
            : base(logger) { }

        public override OpenAIChatPrompt CreatePrompt(MessageContext context)
        {
            var summarizerConfig = ConfigHelper.GetModelConfig(Name);
            OpenAIChatModel model;
{{#useAzureOpenAI}}
            model = new OpenAIChatModel(
                summarizerConfig.Model,
                summarizerConfig.ApiKey,
                new() { Endpoint = new Uri($"{summarizerConfig.Endpoint}/openai/v1") }
            );
{{/useAzureOpenAI}}
{{#useOpenAI}}
            model = new OpenAIChatModel(
                summarizerConfig.ApiKey,
                summarizerConfig.Model
            );
{{/useOpenAI}}

            // Get instructions from SummarizerPrompt
            var instructions = SummarizerPrompt.GetPromptInstructions();
            var prompt = new OpenAIChatPrompt(
                model,
                new ChatPromptOptions().WithInstructions(instructions)
            );

            // Register functions using SummarizerPrompt
            SummarizerPrompt.RegisterFunctions(prompt, context, _logger);

            _logger.LogDebug("✅ Initialized Summarizer Capability!");
            return prompt;
        }

        /// <summary>
        /// Create capability definition for manager registration
        /// </summary>
        public static CapabilityDefinition CreateDefinition(ILogger<SummarizerCapability> logger)
        {
            return new CapabilityDefinition
            {
                Name = "summarizer",
                ManagerDescription =
                    @"**Summarizer**: Use for keywords like:
- 'summarize', 'overview', 'recap', 'conversation history'
- 'what did we discuss', 'catch me up', 'who said what', 'recent messages'",
                Handler = async (context, handlerLogger) =>
                {
                    // Create Summarizer capability (it will create its own ChatModel)
                    var capability = new SummarizerCapability(logger);
                    var result = await capability.ProcessRequestAsync(context);

                    if (!string.IsNullOrEmpty(result.Error))
                    {
                        handlerLogger.LogError(
                            "Error in Summarizer Capability: {Error}",
                            result.Error
                        );
                        return $"Error in Summarizer Capability: {result.Error}";
                    }

                    return result.Response ?? "No response from Summarizer Capability";
                },
            };
        }
    }
}
