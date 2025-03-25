using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using AdaptiveCards;
using Newtonsoft.Json.Linq;

namespace {{SafeProjectName}}.Bot;

public class TeamsMessageExtension : TeamsActivityHandler
{
    private readonly string _searchResultCardFilePath = Path.Combine(".", "Resources", "searchResultCard.json");
    private readonly string _actionAdaptiveCardFilePath = Path.Combine(".", "Resources", "actionCard.json");
    private readonly string _linkUnfurlingAdaptiveCardFilePath = Path.Combine(".", "Resources", "linkUnfurlingCard.json");
    // Action.
    protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
    {
        // The user has chosen to create a card by choosing the 'Create Card' context menu command.
        var actionData = ((JObject)action.Data).ToObject<CardResponse>();
        var templateJson = await System.IO.File.ReadAllTextAsync(_actionAdaptiveCardFilePath, cancellationToken);
        var template = new AdaptiveCards.Templating.AdaptiveCardTemplate(templateJson);
        var adaptiveCardJson = template.Expand(new { title = actionData.Title ?? "", subTitle = actionData.Subtitle ?? "", text = actionData.Text ?? "" });
        var adaptiveCard = AdaptiveCard.FromJson(adaptiveCardJson).Card;
        var attachments = new MessagingExtensionAttachment()
        {
            ContentType = AdaptiveCard.ContentType,
            Content = adaptiveCard
        };

        return new MessagingExtensionActionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = new[] { attachments }
            }
        };
    }

    // Search.
    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        var templateJson = await System.IO.File.ReadAllTextAsync(_searchResultCardFilePath, cancellationToken);
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
                Content = adaptiveCard,
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

    // Link Unfurling.
    // This function can be triggered after this app is installed.
    protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
    {
        var adaptiveCardJson = await System.IO.File.ReadAllTextAsync(_linkUnfurlingAdaptiveCardFilePath, cancellationToken);
        var adaptiveCard = AdaptiveCard.FromJson(adaptiveCardJson).Card;

        var previewCard = new ThumbnailCard
        {
            Title = "Preview Card",
            Text = query.Url,
            Images = new List<CardImage> { new CardImage("https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png") },
        };
        var previewAttachment = new Attachment()
        {
            ContentType = ThumbnailCard.ContentType,
            Content = previewCard
        };

        var attachments = new MessagingExtensionAttachment(AdaptiveCard.ContentType, null, adaptiveCard, preview: previewAttachment);

        // By default the link unfurling result is cached in Teams for 30 minutes.
        // The code has set a cache policy and removed the cache for the app. Learn more here: https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/how-to/link-unfurling?tabs=dotnet%2Cadvantages#remove-link-unfurling-cache
        var action = new CardAction
        {
            Type = "setCachePolicy",
            Value = "{\"type\":\"no-cache\"}",
        };

        var suggestedActions = new MessagingExtensionSuggestedAction([action]);
        var result = new MessagingExtensionResult("list", "result", new[] { attachments }, suggestedActions);
        return new MessagingExtensionResponse(result);
    }


    internal class CardResponse
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Text { get; set; }
    }
}