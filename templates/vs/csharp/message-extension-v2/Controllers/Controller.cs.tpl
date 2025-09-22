using System.Text.Json;

using Microsoft.Teams.Api.Cards;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities.Invokes;
using Microsoft.Teams.Apps.Annotations;
using Microsoft.Teams.Cards;

namespace {{SafeProjectName}}.Controllers
{
    [TeamsController]
    public class Controller()
    {
        [MessageExtension.Query]
        public Microsoft.Teams.Api.MessageExtensions.Response OnMessageExtensionQuery(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.QueryActivity activity,
            [Context] IContext.Client client,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_QUERY] Search query received");

            var commandId = activity.Value?.CommandId;
            var query = activity.Value?.Parameters?.FirstOrDefault(p => p.Name == "searchQuery")?.Value?.ToString() ?? "";

            log.Info($"[MESSAGE_EXT_QUERY] Command: {commandId}, Query: {query}");

            if (commandId == "searchQuery")
            {
                return MessageExtensionResponseHelper.CreateSearchResults(query, log);
            }

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Result,
                    AttachmentLayout = Microsoft.Teams.Api.Attachment.Layout.List,
                    Attachments = new List<Microsoft.Teams.Api.MessageExtensions.Attachment>()
                }
            };
        }

        [MessageExtension.SubmitAction]
        public Microsoft.Teams.Api.MessageExtensions.Response OnMessageExtensionSubmit(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.SubmitActionActivity activity,
            [Context] IContext.Client client,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_SUBMIT] Action submit received");

            var commandId = activity.Value?.CommandId;
            var data = activity.Value?.Data as JsonElement?;

            log.Info($"[MESSAGE_EXT_SUBMIT] Command: {commandId}");
            log.Info($"[MESSAGE_EXT_SUBMIT] Data: {JsonSerializer.Serialize(data)}");

            switch (commandId)
            {
                case "createCard":
                    return MessageExtensionResponseHelper.HandleCreateCard(data, log);

                case "getMessageDetails":
                    return MessageExtensionResponseHelper.HandleGetMessageDetails(activity, log);

                default:
                    log.Error($"[MESSAGE_EXT_SUBMIT] Unknown command: {commandId}");
                    return MessageExtensionResponseHelper.CreateErrorActionResponse("Unknown command");
            }
        }

        [MessageExtension.QueryLink]
        public Microsoft.Teams.Api.MessageExtensions.Response OnMessageExtensionQueryLink(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.QueryLinkActivity activity,
            [Context] IContext.Client client,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_QUERY_LINK] Link unfurling received");

            var url = activity.Value?.Url;
            log.Info($"[MESSAGE_EXT_QUERY_LINK] URL: {url}");

            if (string.IsNullOrEmpty(url))
            {
                return MessageExtensionResponseHelper.CreateErrorResponse("No URL provided");
            }

            return MessageExtensionResponseHelper.CreateLinkUnfurlResponse(url, log);
        }

        [MessageExtension.SelectItem]
        public Microsoft.Teams.Api.MessageExtensions.Response OnMessageExtensionSelectItem(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.SelectItemActivity activity,
            [Context] IContext.Client client,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_SELECT_ITEM] Item selection received");

            var selectedItem = activity.Value;
            log.Info($"[MESSAGE_EXT_SELECT_ITEM] Selected: {JsonSerializer.Serialize(selectedItem)}");

            return MessageExtensionResponseHelper.CreateItemSelectionResponse(selectedItem, log);
        }
    }

}
