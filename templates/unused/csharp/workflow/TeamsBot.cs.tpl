using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using {{SafeProjectName}}.CardActions;
using {{SafeProjectName}}.Commands;

namespace {{SafeProjectName}}
{
    /// <summary>
    /// Bot handler.
    /// You can add your customization code here to extend your bot logic if needed.
    /// </summary>
    public class TeamsBot : AgentApplication
    {
        private readonly GenericCommandHandler _genericCommandHandler;
        private readonly HelloWorldCommandHandler _helloWorldCommandHandler;
        private readonly ILogger<TeamsBot> _logger;
        
        public TeamsBot(AgentApplicationOptions options, 
                       GenericCommandHandler genericCommandHandler,
                       HelloWorldCommandHandler helloWorldCommandHandler,
                       ILogger<TeamsBot> logger) : base(options)
        {      
            _genericCommandHandler = genericCommandHandler;
            _helloWorldCommandHandler = helloWorldCommandHandler;
            _logger = logger;
            
            OnConversationUpdate(ConversationUpdateEvents.MembersAdded, OnMembersAddedAsync);
            
            // Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
            OnActivity(ActivityTypes.Message, OnMessageReceivedAsync);

            AdaptiveCards.OnActionExecute(DoStuffActionHandler.TriggerVerb, DoStuffActionHandler.handler);
        }

        protected async Task OnMembersAddedAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
        {
            var welcomeText = "Welcome to the Workflow Bot! I can help you work through the Adaptive Card and perform various tasks. Type \"helloworld\" or \"help\" to get started.";
            foreach (var member in turnContext.Activity.MembersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText), cancellationToken);
                }
            }
        }
        
        protected async Task OnMessageReceivedAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
        {
            var message = turnContext.Activity.Text?.Trim();
            
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            _logger?.LogInformation($"Processing message: {message}");

            IActivity response = null;

            // Check HelloWorld handler patterns first (more specific)
            foreach (var pattern in _helloWorldCommandHandler.TriggerPatterns)
            {
                if (pattern.IsMatch(message))
                {
                    _logger?.LogInformation($"Message matched HelloWorld pattern: {pattern}");
                    response = await _helloWorldCommandHandler.HandleCommandAsync(turnContext, cancellationToken);
                    break;
                }
            }

            // If no HelloWorld pattern matched, check Generic handler patterns
            if (response == null)
            {
                foreach (var pattern in _genericCommandHandler.TriggerPatterns)
                {
                    if (pattern.IsMatch(message))
                    {
                        _logger?.LogInformation($"Message matched Generic pattern: {pattern}");
                        response = await _genericCommandHandler.HandleCommandAsync(turnContext, cancellationToken);
                        break;
                    }
                }
            }

            // Send the response if we got one
            if (response != null)
            {
                await turnContext.SendActivityAsync(response, cancellationToken);
            }
            else
            {
                _logger?.LogWarning($"No handler matched for message: {message}");
            }
        }
    }
}
