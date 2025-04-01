namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.BotBuilder;
    using Microsoft.TeamsFx.Conversation;
    using System.Threading;
    using System.Threading.Tasks;

    public class TestCommandHandler : ITeamsCommandHandler
    {
        public bool IsTriggered { get; set; } = false;

        public CommandMessage ReceivedMessage { get; private set; }

        public IEnumerable<ITriggerPattern> TriggerPatterns { get; }

        public TestCommandHandler(string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
            {
                throw new ArgumentNullException(nameof(pattern));
            }

            TriggerPatterns = new List<ITriggerPattern> { new StringTrigger(pattern) };
            IsTriggered = false;
        }

        public TestCommandHandler(IEnumerable<ITriggerPattern> triggerPatterns)
        {
            TriggerPatterns = triggerPatterns;
            IsTriggered = false;
        }

        public Task<ICommandResponse> HandleCommandAsync(ITurnContext turnContext, CommandMessage message, CancellationToken cancellationToken = default)
        {
            IsTriggered = true;
            ReceivedMessage = message;

            var Response = new TextCommandResponse("This is an sample response from command bot!");
            return Task.FromResult<ICommandResponse>(Response);
        }
    }
}
