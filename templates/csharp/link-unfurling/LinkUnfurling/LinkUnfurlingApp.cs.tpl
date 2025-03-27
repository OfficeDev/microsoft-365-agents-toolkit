using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Extensions.Teams.Compat;
using Microsoft.Agents.Extensions.Teams.Models;
using AdaptiveCards;

namespace {{SafeProjectName}}.LinkUnfurling;

public class LinkUnfurlingApp : TeamsActivityHandler
{ 
    private readonly string _adaptiveCardFilePath = Path.Combine(".", "Resources", "helloWorldCard.json");


    // Link Unfurling.
    // This function can be triggered after this app is installed.
    protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
    {
        var adaptiveCardJson = await System.IO.File.ReadAllTextAsync(_adaptiveCardFilePath, cancellationToken);
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

        var attachments = new MessagingExtensionAttachment(AdaptiveCard.ContentType, null, adaptiveCard.ToJson(), preview: previewAttachment);

        // By default the link unfurling result is cached in Teams for 30 minutes.
        // The code has set a cache policy and removed the cache for the app. Learn more here: https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/how-to/link-unfurling?tabs=dotnet%2Cadvantages#remove-link-unfurling-cache
        var action = new CardAction
        {
            Type = "setCachePolicy",
            Value = "{\"type\":\"no-cache\"}",
        };
        
        var suggestedActions = new MessagingExtensionSuggestedAction(new[] { action });
        var result = new MessagingExtensionResult("list", "result", new[] { attachments }, suggestedActions);
        return new MessagingExtensionResponse(result);
    }

    // Zero Install Link Unfurling
    // This function can be triggered if this app sets "supportsAnonymizedPayloads": true in manifest and is uploaded to org's app catalog.
    protected override async Task<MessagingExtensionResponse> OnTeamsAnonymousAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
    {
        var adaptiveCardJson = await System.IO.File.ReadAllTextAsync(_adaptiveCardFilePath, cancellationToken);
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

        var attachments = new MessagingExtensionAttachment(AdaptiveCard.ContentType, null, adaptiveCard.ToJson(), preview: previewAttachment);

        // By default the link unfurling result is cached in Teams for 30 minutes.
        // The code has set a cache policy and removed the cache for the app. Learn more here: https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/how-to/link-unfurling?tabs=dotnet%2Cadvantages#remove-link-unfurling-cache
        var action = new CardAction
        {
            Type = "setCachePolicy",
            Value = "{\"type\":\"no-cache\"}",
        };

        var suggestedActions = new MessagingExtensionSuggestedAction(new[] { action });
        var result = new MessagingExtensionResult("list", "result", new[] { attachments }, suggestedActions);
        return new MessagingExtensionResponse(result);
    }
}

