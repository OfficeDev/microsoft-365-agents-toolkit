using AdaptiveCards;
using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Extensions.Teams.Compat;
using Microsoft.Agents.Extensions.Teams.Models;
using Newtonsoft.Json.Linq;

namespace {{SafeProjectName}}.Search;

public class SearchApp : TeamsActivityHandler
{
    private readonly string _adaptiveCardFilePath = Path.Combine(".", "Resources", "helloWorldCard.json");
    // Search
    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        var templateJson = await System.IO.File.ReadAllTextAsync(_adaptiveCardFilePath, cancellationToken);
        var template = new AdaptiveCards.Templating.AdaptiveCardTemplate(templateJson);

        var text = query?.Parameters?[0]?.Value as string ?? string.Empty;
        var packages = await FindPackages(text);
        // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
        var attachments = packages.Select(package =>
        {
            var previewCard = new ThumbnailCard { Title = package.Item1 };

            var adaptiveCardJson = template.Expand(new { name = package.Item1, description = package.Item3 });
            var adaptiveCard = AdaptiveCard.FromJson(adaptiveCardJson).Card;
            if (!string.IsNullOrEmpty(package.Item5))
            {
                previewCard.Images = new List<CardImage>() { new CardImage(package.Item5, "Icon") };
                adaptiveCard.Body.Insert(0, new AdaptiveImage()
                {
                    Url = new Uri(package.Item5),
                    Style = AdaptiveImageStyle.Person,
                    Size = AdaptiveImageSize.Small,
                });
            }
            var attachment = new MessagingExtensionAttachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard.ToJson(),
                Preview = previewCard.ToAttachment()
            };

            return attachment;
        }).ToList();

        return new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = attachments
            }
        };
    }

    // Generate a set of substrings to illustrate the idea of a set of results coming back from a query. 
    private async Task<IEnumerable<(string, string, string, string, string)>> FindPackages(string text)
    {
        var httpClient = new HttpClient();
        var response = await httpClient.GetStringAsync($"https://azuresearch-usnc.nuget.org/query?q=id:{text}&prerelease=true");
        var obj = JObject.Parse(response);
        return obj["data"].Select(item => (
            item["id"].ToString(),
            item["version"].ToString(),
            item["description"].ToString(),
            item["projectUrl"]?.ToString(),
            item["iconUrl"]?.ToString()));
    }
}