using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Agent
{
    public class ManagerResult
    {
        public string Response { get; set; } = string.Empty;
    }

    public class Manager
    {
        private OpenAIChatPrompt? _prompt;
        private readonly ILogger<Manager> _logger;
        private readonly List<CapabilityDefinition> _capabilityDefinitions;
        private bool _isInitialized = false;

        public Manager(
            ILoggerFactory loggerFactory,
            List<CapabilityDefinition> capabilityDefinitions
        )
        {
            _logger = loggerFactory.CreateLogger<Manager>();
            _capabilityDefinitions = capabilityDefinitions;
            _logger.LogDebug(
                "🎯 Manager instance created with {Count} capabilities",
                capabilityDefinitions.Count
            );
        }

        /// <summary>
        /// Get capabilities list (for prompt registration)
        /// </summary>
        public List<CapabilityDefinition> GetCapabilities() => _capabilityDefinitions;

        /// <summary>
        /// Generate manager prompt instructions by filling in the template
        /// </summary>
        public string GenerateManagerPrompt(List<CapabilityDefinition> capabilities)
        {
            // Create numbered list of capability names
            var capabilityList = string.Join(
                "\n",
                capabilities.Select((cap, i) => $"{i + 1}. **{cap.Name}**")
            );

            // Create capability descriptions
            var capabilityDescriptions = string.Join(
                "\n",
                capabilities.Select(cap => cap.ManagerDescription)
            );

            // Get template and replace placeholders
            var template = ManagerPrompt.GetPromptTemplate();
            var instructions = template
                .Replace("{capabilityList}", capabilityList)
                .Replace("{capabilityDescriptions}", capabilityDescriptions);

            return instructions;
        }

        private async Task<OpenAIChatPrompt> CreateManagerPrompt(MessageContext context)
        {
            _logger.LogDebug(
                "📋 Creating manager prompt with {Count} capabilities",
                _capabilityDefinitions.Count
            );

            var instructions = GenerateManagerPrompt(_capabilityDefinitions);

            var promptOptions = new ChatPromptOptions();
            promptOptions.WithInstructions(instructions);

            // Create Manager's own ChatModel (using lighter model)
            var managerConfig = ConfigHelper.GetModelConfig("manager");
            OpenAIChatModel chatModel;
{{#useAzureOpenAI}}
            chatModel = new OpenAIChatModel(
                managerConfig.Model,
                managerConfig.ApiKey,
                new() { Endpoint = new Uri($"{managerConfig.Endpoint}/openai/v1") }
            );
{{/useAzureOpenAI}}
{{#useOpenAI}}
            chatModel = new OpenAIChatModel(
                managerConfig.ApiKey,
                managerConfig.Model
            );
{{/useOpenAI}}
            _logger.LogDebug("✅ Manager chat model created: {Model}", managerConfig.Model);

            var prompt = new OpenAIChatPrompt(chatModel, promptOptions);

            // Register all functions using ManagerPrompt
            ManagerPrompt.RegisterFunctions(prompt, context, this, _logger);
            ManagerPrompt.RegisterCapabilityFunctions(prompt, context, this, _logger);

            _logger.LogDebug("✅ Manager prompt created with all functions");
            return prompt;
        }

        /// <summary>
        /// Calculate time range from natural language
        /// </summary>
        public void CalculateTimeRange(MessageContext context, string timePhrase)
        {
            _logger.LogDebug($"🕒 Calculating time range for: \"{timePhrase}\"");

            var timeRange = TimeRangeUtils.ExtractTimeRange(timePhrase, _logger);

            if (timeRange != null)
            {
                context.StartTime = timeRange.Value.From.ToString("o");
                context.EndTime = timeRange.Value.To.ToString("o");
                _logger.LogInformation(
                    $"✅ Time range calculated: {context.StartTime} to {context.EndTime}"
                );
            }
            else
            {
                _logger.LogWarning($"Could not parse time phrase: \"{timePhrase}\"");
            }
        }

        /// <summary>
        /// Clear conversation history
        /// </summary>
        public async Task ClearConversationHistoryAsync(MessageContext context)
        {
            _logger.LogDebug("Clearing conversation history");
            await context.Memory.ClearAsync();
            _logger.LogDebug("The conversation history has been cleared!");
        }

        /// <summary>
        /// Delegate to a specific capability
        /// </summary>
        public async Task<string> DelegateToCapability(
            string capabilityName,
            MessageContext context
        )
        {
            var capability = _capabilityDefinitions.FirstOrDefault(c =>
                c.Name.Equals(capabilityName, StringComparison.OrdinalIgnoreCase)
            );

            if (capability == null)
            {
                _logger.LogWarning($"Capability '{capabilityName}' not found");
                return $"Capability '{capabilityName}' is not available.";
            }

            _logger.LogInformation($"🚀 Delegating to capability: {capabilityName}");
            _logger.LogDebug($"📝 User request: \"{context.Text}\"");

            var result = await capability.Handler(context, _logger);

            _logger.LogInformation($"✅ Capability {capabilityName} returned result");
            return result;
        }

        private async Task InitializeAsync(MessageContext context)
        {
            if (!_isInitialized)
            {
                _logger.LogDebug("🔧 Initializing Manager");
                _prompt = await CreateManagerPrompt(context);
                _isInitialized = true;
                _logger.LogInformation("✅ Manager fully initialized");
            }
        }

        public async Task<ManagerResult> ProcessRequestAsync(MessageContext context)
        {
            try
            {
                _logger.LogInformation(
                    "📨 Processing request from user: {UserId}, text: \"{Text}\"",
                    context.UserId ?? "unknown",
                    context.Text
                );

                await InitializeAsync(context);

                if (_prompt == null)
                {
                    throw new InvalidOperationException("Manager prompt is not initialized");
                }

                _logger.LogDebug("🤖 Sending request to LLM: {Text}", context.Text);

                // Manager doesn't need conversation history
                // Let the LLM decide if it needs to call summarizer (which has access to memory)
                var response = await _prompt.Send(
                    context.Text,
                    CancellationToken.None
                );

                _logger.LogInformation(
                    "✅ LLM response received, content length: {Length} chars",
                    response.Content?.Length ?? 0
                );

                return new ManagerResult { Response = response.Content ?? "No response generated" };
            }
            catch (Exception error)
            {
                _logger.LogError(error, "❌ Error in Manager");
                return new ManagerResult
                {
                    Response =
                        $"Sorry, I encountered an error processing your request: {error.Message}",
                };
            }
        }
    }
}
