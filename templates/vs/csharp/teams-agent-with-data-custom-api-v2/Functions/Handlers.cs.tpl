using AdaptiveCards.Templating;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Api.Activities;
using OpenAPIClient;
using RestSharp;

namespace {{SafeProjectName}}.Functions
{
    public class Handlers(APIClient Client, IContext.Accessor accessor)
    {
        private IContext<IActivity> context => accessor.Value!;

        // Replace with function handler code

        private static RequestParams ParseRequestParams(IDictionary<string, object?> args)
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

        private static string RenderCardToMessage(string cardTemplatePath, string data)
        {
            try
            {
                var templateString = File.ReadAllText(cardTemplatePath);
                AdaptiveCardTemplate template = new AdaptiveCardTemplate(templateString);
                return template.Expand(data);
               

            }
            catch (Exception ex)
            {
                throw new Exception("Failed to render adaptive card: " + ex.Message);
            }
        }
    }
}
