using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.ActionItems
{
    /// <summary>
    /// Action Items capability for extracting action items from conversations
    /// </summary>
    public class ActionItemsCapability : BaseCapability
    {
        public override string Name => "action_items";

        public ActionItemsCapability(ILogger logger)
            : base(logger) { }

        public override OpenAIChatPrompt CreatePrompt(MessageContext context)
        {
            var actionItemsConfig = ConfigHelper.GetModelConfig("actionItems");
            OpenAIChatModel model;
{{#useAzureOpenAI}}
            model = new OpenAIChatModel(
                actionItemsConfig.Model,
                actionItemsConfig.ApiKey,
                new() { Endpoint = new Uri($"{actionItemsConfig.Endpoint}/openai/v1") }
            );
{{/useAzureOpenAI}}
{{#useOpenAI}}
            model = new OpenAIChatModel(
                actionItemsConfig.ApiKey,
                actionItemsConfig.Model
            );
{{/useOpenAI}}

            // Get instructions from ActionItemsPrompt
            var instructions = ActionItemsPrompt.GetPromptInstructions();
            var prompt = new OpenAIChatPrompt(
                model,
                new ChatPromptOptions().WithInstructions(instructions)
            );

            // Register functions using ActionItemsPrompt
            ActionItemsPrompt.RegisterFunctions(prompt, context, _logger);

            _logger.LogDebug(
                "✅ Initialized Action Items Capability using {Count} members from context",
                context.Members.Count
            );
            return prompt;
        }

        /// <summary>
        /// Create capability definition for manager registration
        /// </summary>
        public static CapabilityDefinition CreateDefinition(ILogger<ActionItemsCapability> logger)
        {
            return new CapabilityDefinition
            {
                Name = "action_items",
                ManagerDescription =
                    @"**Action Items**: Use for requests like:
- ""next steps"", ""to-do"", ""assign task"", ""my tasks"", ""what needs to be done""",
                Handler = async (context, handlerLogger) =>
                {
                    // Create Action Items capability (it will create its own ChatModel)
                    var capability = new ActionItemsCapability(logger);
                    var result = await capability.ProcessRequestAsync(context);

                    if (!string.IsNullOrEmpty(result.Error))
                    {
                        handlerLogger.LogError(
                            "Error in Action Items Capability: {Error}",
                            result.Error
                        );
                        return $"Error in Action Items Capability: {result.Error}";
                    }

                    return result.Response ?? "No response from Action Items Capability";
                },
            };
        }
    }
}
