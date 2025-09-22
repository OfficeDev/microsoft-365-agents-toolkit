using AdaptiveCards.Templating;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.AI.Action;
using Microsoft.Teams.AI.AI;
using Microsoft.Teams.AI.State;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Schema;
using RestSharp;
using OpenAPIClient;

namespace {{SafeProjectName}}
{
    public class APIActions
    {
        private APIClient Client;

        public APIActions()
        {
            Client = new APIClient("{{OPENAPI_SPEC_PATH}}");
        }

        // Replace with action code

        private static IMessageActivity RenderCardToMessage(string cardTemplatePath, string data)
        {
            try
            {
                var templateString = File.ReadAllText(cardTemplatePath);
                AdaptiveCardTemplate template = new AdaptiveCardTemplate(templateString);
                var cardBody = template.Expand(data);

                Attachment attachment = new Attachment()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = JObject.Parse(cardBody)
                };

                return MessageFactory.Attachment(attachment);
            }
            catch (Exception ex) { 
                throw new Exception("Failed to render adaptive card: " +  ex.Message);
            }
        }

        private static RequestParams ParseRequestParams(Dictionary<string, object> args)
        {
            RequestParams requestParam = new RequestParams
            {
                PathObject = args.ContainsKey("path") ? args["path"] : null,
                HeaderObject = args.ContainsKey("header") ? args["header"] : null,
                QueryObject = args.ContainsKey("query") ? args["query"] : null,
                RequestBody = args.ContainsKey("body") ? args["body"] : null
            };
            return requestParam;
        }

        [Action(AIConstants.UnknownActionName)]
        public async Task<string> UnknownActionAsync([ActionTurnContext] TurnContext turnContext, [ActionName] string action)
        {
            await turnContext.SendActivityAsync(MessageFactory.Text("Unable to find a matching API to call"));
            return "unknown action";
        }
    }
}
