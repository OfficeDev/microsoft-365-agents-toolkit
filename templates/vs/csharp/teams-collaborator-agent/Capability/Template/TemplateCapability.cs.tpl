using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability.Template
{
    /// <summary>
    /// Template capability - Use this as a starting point for creating new capabilities
    /// 
    /// TEMPLATE INSTRUCTIONS:
    /// 1. Copy this entire Template folder and rename it to your capability name (e.g., "MyCapability")
    /// 2. Rename the class from TemplateCapability to YourCapabilityName (e.g., MyCapability)
    /// 3. Update the Name property to return your capability's identifier (lowercase with underscores)
    /// 4. Customize TemplatePrompt.cs with your capability's logic
    /// 5. Add your capability to Program.cs (see instructions there)
    /// 
    /// TIPS:
    /// - Keep capability names short and descriptive (e.g., "summarizer", "action_items", "search")
    /// - Use the logger to debug and track what your capability is doing
    /// - Test with simple scenarios first before adding complexity
    /// </summary>
    public class TemplateCapability : BaseCapability
    {
        // TODO: Change this to your capability's name (lowercase, use underscores for spaces)
        public override string Name => "template";

        public TemplateCapability(ILogger logger)
            : base(logger) { }

        public override OpenAIChatPrompt CreatePrompt(MessageContext context)
        {
            // Get model configuration for your capability
            // TODO: Update "template" to match your capability name in appsettings.json
            var templateConfig = ConfigHelper.GetModelConfig("template");
            
            OpenAIChatModel model;
{{#useAzureOpenAI}}
            model = new OpenAIChatModel(
                templateConfig.Model,
                templateConfig.ApiKey,
                new() { Endpoint = new Uri($"{templateConfig.Endpoint}/openai/v1") }
            );
{{/useAzureOpenAI}}
{{#useOpenAI}}
            model = new OpenAIChatModel(
                templateConfig.ApiKey,
                templateConfig.Model
            );
{{/useOpenAI}}

            // Get instructions from TemplatePrompt
            var instructions = TemplatePrompt.GetPromptInstructions();
            var prompt = new OpenAIChatPrompt(
                model,
                new ChatPromptOptions().WithInstructions(instructions)
            );

            // Register functions using TemplatePrompt
            TemplatePrompt.RegisterFunctions(prompt, context, _logger);

            _logger.LogDebug("✅ Initialized Template Capability!");
            return prompt;
        }

        /// <summary>
        /// Create capability definition for manager registration
        /// 
        /// This method is called from Program.cs to register your capability.
        /// The ManagerDescription tells the Manager when to route requests to your capability.
        /// </summary>
        public static CapabilityDefinition CreateDefinition(ILogger<TemplateCapability> logger)
        {
            return new CapabilityDefinition
            {
                // TODO: Update to match your capability name
                Name = "template",
                
                // TODO: Describe when this capability should be used
                // The Manager uses this to decide which capability handles each request
                ManagerDescription =
                    @"**Template**: Use for requests like:
- TODO: Add keywords and phrases that indicate this capability should be used
- TODO: For example: ""template"", ""example"", ""demo""
- TODO: Be specific so the Manager can route requests correctly",
                
                // Handler that processes requests for this capability
                Handler = async (context, handlerLogger) =>
                {
                    // Create Template capability (it will create its own ChatModel)
                    var capability = new TemplateCapability(logger);
                    var result = await capability.ProcessRequestAsync(context);

                    if (!string.IsNullOrEmpty(result.Error))
                    {
                        handlerLogger.LogError(
                            "❌ Error in Template Capability: {Error}",
                            result.Error
                        );
                        return $"Error in Template Capability: {result.Error}";
                    }

                    return result.Response ?? "No response from Template Capability";
                },
            };
        }
    }
}
