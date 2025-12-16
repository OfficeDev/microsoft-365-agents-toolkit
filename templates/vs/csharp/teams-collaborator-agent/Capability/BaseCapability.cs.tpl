using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.AI.Prompts;

namespace {{SafeProjectName}}.Capability
{
    public class CapabilityDefinition
    {
        public string Name { get; set; } = string.Empty;
        public string ManagerDescription { get; set; } = string.Empty;
        public Func<MessageContext, ILogger, Task<string>> Handler { get; set; } = null!;
    }

    public class CapabilityResult
    {
        public string Response { get; set; } = string.Empty;
        public string? Error { get; set; }
    }

    public interface ICapability
    {
        string Name { get; }
        OpenAIChatPrompt CreatePrompt(MessageContext context);
        Task<CapabilityResult> ProcessRequestAsync(MessageContext context);
    }

    public abstract class BaseCapability : ICapability
    {
        protected readonly ILogger _logger;

        public abstract string Name { get; }

        protected BaseCapability(ILogger logger) => _logger = logger;

        public abstract OpenAIChatPrompt CreatePrompt(MessageContext context);

        public virtual async Task<CapabilityResult> ProcessRequestAsync(MessageContext context)
        {
            try
            {
                var prompt = CreatePrompt(context);
                var response = await prompt.Send(
                    context.Text,
                    CancellationToken.None
                );

                return new CapabilityResult
                {
                    Response = response.Content ?? "No response generated",
                };
            }
            catch (Exception error)
            {
                _logger.LogError(error, $"Error in {Name} capability");
                return new CapabilityResult { Response = string.Empty, Error = error.Message };
            }
        }
    }
}
