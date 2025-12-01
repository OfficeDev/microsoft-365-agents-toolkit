using {{SafeProjectName}}.Models;
using AdaptiveCards.Templating;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App.AdaptiveCards;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;

namespace {{SafeProjectName}}.CardActions
{
    public class DoStuffActionHandler
    {
        private readonly static string _responseCardFilePath = Path.Combine(".", "Resources", "DoStuffActionResponse.json");

        /// <summary>
        /// A global unique string associated with the `OnActionExecute` action.
        /// The value should be the same as the `verb` property which you define in your adaptive card JSON.
        /// </summary>
        public static string TriggerVerb => "doStuff";

        public static ActionExecuteHandler handler = HandleActionInvokedAsync;

        private static async Task<AdaptiveCardInvokeResponse> HandleActionInvokedAsync(ITurnContext turnContext, ITurnState state, object cardData, CancellationToken cancellationToken = default)
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
            return  AdaptiveCardInvokeResponseFactory.AdaptiveCard(cardContent);

            /**
             * If you want to send invoke response with text message, you can:
             *
             * return AdaptiveCardInvokeResponseFactory.Message("[ACK] Successfully!");
            */

            /**
             * If you want to send invoke response with error message, you can:
             *
             * return AdaptiveCardInvokeResponseFactory.Error(HttpStatusCode.BadRequest, "400","The incoming request is invalid.");
             */
        }
    }
}
