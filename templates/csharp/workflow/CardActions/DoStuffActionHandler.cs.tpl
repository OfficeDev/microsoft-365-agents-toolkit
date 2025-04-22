using {{SafeProjectName}}.Models;
using AdaptiveCards.Templating;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Core.Models;
using Microsoft.TeamsFx.Conversation;

namespace {{SafeProjectName}}.CardActions
{
    public class DoStuffActionHandler : IAdaptiveCardActionHandler
    {
        private readonly string _responseCardFilePath = Path.Combine(".", "Resources", "DoStuffActionResponse.json");

        /// <summary>
        /// A global unique string associated with the `Action.Execute` action.
        /// The value should be the same as the `verb` property which you define in your adaptive card JSON.
        /// </summary>
        public string TriggerVerb => "doStuff";

        /// <summary>
        /// Indicate how your action response card is sent in the conversation.
        /// By default, the response card can only be updated for the interactor who trigger the action.
        /// </summary>
        public AdaptiveCardResponse AdaptiveCardResponse => AdaptiveCardResponse.ReplaceForInteractor;

        public async Task<InvokeResponse> HandleActionInvokedAsync(ITurnContext turnContext, object cardData, CancellationToken cancellationToken = default)
        {
            // Read adaptive card template
            var cardTemplate = await File.ReadAllTextAsync(_responseCardFilePath, cancellationToken);

            // Render adaptive card content
            var cardContent = new AdaptiveCardTemplate(cardTemplate).Expand
            (
                new HelloWorldModel
                {
                    Title = "Hello World Bot",
                    Body = $"Congratulations! Your {TriggerVerb} action is processed successfully!",
                }
            );

            // Send invoke response with adaptive card
            return InvokeResponseFactory.AdaptiveCard(cardContent);

            /**
             * If you want to send invoke response with text message, you can:
             *
             * return InvokeResponseFactory.TextMessage("[ACK] Successfully!");
            */

            /**
             * If you want to send invoke response with error message, you can:
             *
             * return InvokeResponseFactory.ErrorResponse(InvokeResponseErrorCode.BadRequest, "The incoming request is invalid.");
             */
        }
    }
}
