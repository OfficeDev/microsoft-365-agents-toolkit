using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Core.Models;
using Microsoft.TeamsFx.Conversation;

namespace {{SafeProjectName}}.Commands
{
    /// <summary>
    /// The <see cref="GenericCommandHandler"/> registers patterns with the <see cref="ITeamsCommandHandler"/> and
    /// responds with appropriate messages if the user types general command inputs, such as "hi", "hello", and "help".
    /// </summary>
    public class GenericCommandHandler : ITeamsCommandHandler
    {
        private readonly ILogger<GenericCommandHandler> _logger;
    
        public IEnumerable<ITriggerPattern> TriggerPatterns => new List<ITriggerPattern>
        {
            // Used to trigger the command handler when the user enters a generic or unknown command
            new RegExpTrigger("^.+$")
        };

        public GenericCommandHandler(ILogger<GenericCommandHandler> logger)
        {
            _logger = logger;
        }

        public async Task<ICommandResponse> HandleCommandAsync(ITurnContext turnContext, CommandMessage message, CancellationToken cancellationToken = default)
        {
            _logger?.LogInformation($"App received message: {message.Text}");
    
            // Determine the appropriate response based on the command  
            string responseText;
            switch (message.Text)
            {
                case "hi":
                    responseText = "Hi there! I'm your Workflow Bot, here to assist you with your tasks. Type 'help' for a list of available commands.";
                    break;
                case "hello":
                    responseText = "Hello! I'm your Workflow Bot, always ready to help you out. If you need assistance, just type 'help' to see the available commands.";
                    break;
                case "help":
                    responseText = "Here's a list of commands I can help you with:\n" +
                                   "- 'hi' or 'hello': Say hi or hello to me, and I'll greet you back.\n" +
                                   "- 'help': Get a list of available commands.\n" +
                                   "- 'helloworld': See a sample workflow from me.\n" +
                                   "\nFeel free to ask for help anytime you need it!";
                    break;
                default:
                    responseText = "Sorry, command unknown. Please type 'help' to see the list of available commands.";
                    break;
            }

            // Build the response activity using MessageFactory from Microsoft.Agents.Core.Models
            var activity = MessageFactory.Text(responseText);

            // Send response  
            return new ActivityCommandResponse(activity);
        }
    }
}