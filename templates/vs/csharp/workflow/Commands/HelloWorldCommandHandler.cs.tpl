using {{SafeProjectName}}.Models;
using AdaptiveCards.Templating;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Core.Models;
using System.Text.RegularExpressions;

namespace {{SafeProjectName}}.Commands
{
    /// <summary>
    /// The <see cref="HelloWorldCommandHandler"/> registers a pattern with the <see cref="ITeamsCommandHandler"/> and
    /// responds with an Adaptive Card if the user types the <see cref="TriggerPatterns"/>.
    /// </summary>
    public class HelloWorldCommandHandler
    {
        private readonly ILogger<HelloWorldCommandHandler> _logger;
        private readonly string _adaptiveCardFilePath = Path.Combine(".", "Resources", "HelloWorldCommandResponse.json");

        public IEnumerable<Regex> TriggerPatterns => new List<Regex>
        {
            // Used to trigger the command handler if the command text contains 'helloWorld'
            new Regex("helloWorld", RegexOptions.IgnoreCase)
        };

        public HelloWorldCommandHandler(ILogger<HelloWorldCommandHandler> logger)
        {
            _logger = logger;
        }

        public async Task<IActivity> HandleCommandAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            _logger?.LogInformation($"App received message: {turnContext.Activity.Text}");

            // Read adaptive card template
            var cardTemplate = await File.ReadAllTextAsync(_adaptiveCardFilePath, cancellationToken);

            // Render adaptive card content
            var cardContent = new AdaptiveCardTemplate(cardTemplate).Expand
            (
                new HelloWorldModel
                {
                    Title = "Your Hello World Bot is Running",
                    Body = "Congratulations! Your hello world bot is running. Click the button below to trigger an action.",
                }
            );

            // Build attachment
            var activity = MessageFactory.Attachment
            (
                new Attachment
                {
                    ContentType = "application/vnd.microsoft.card.adaptive",
                    Content = cardContent
                }
            );

            // send response
            return activity;
        }
    }
}
