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
                return CreateSearchResults(query, log);
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
                    return HandleCreateCard(data, log);

                case "getMessageDetails":
                    return HandleGetMessageDetails(activity, log);

                default:
                    log.Error($"[MESSAGE_EXT_SUBMIT] Unknown command: {commandId}");
                    return CreateErrorActionResponse("Unknown command");
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
                return CreateErrorResponse("No URL provided");
            }

            return CreateLinkUnfurlResponse(url, log);
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

            return CreateItemSelectionResponse(selectedItem, log);
        }

        [MessageExtension.QuerySettingsUrl]
        public Microsoft.Teams.Api.MessageExtensions.Response OnMessageExtensionQuerySettingsUrl(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.QuerySettingUrlActivity activity,
            [Context] IContext.Client client,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_QUERY_SETTINGS_URL] Settings URL requested");

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Config,
                    Text = "Settings configuration would be handled here"
                }
            };
        }

        [MessageExtension.FetchTask]
        public async Task<Microsoft.Teams.Api.MessageExtensions.ActionResponse> OnMessageExtensionFetchTask(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.FetchTaskActivity activity,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_FETCH_TASK] Fetch task received");

            var commandId = activity.Value?.CommandId;
            log.Info($"[MESSAGE_EXT_FETCH_TASK] Command: {commandId}");

            return CreateFetchTaskResponse(commandId, log);
        }

        [MessageExtension.Setting]
        public Microsoft.Teams.Api.MessageExtensions.Response OnMessageExtensionSetting(
            [Context] Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.SettingActivity activity,
            [Context] IContext.Client client,
            [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("[MESSAGE_EXT_SETTING] Settings received");

            var state = activity.Value?.State;
            log.Info($"[MESSAGE_EXT_SETTING] State: {state}");

            if (state == "cancel")
            {
                log.Info("[MESSAGE_EXT_SETTING] Settings cancelled by user");
                return new Microsoft.Teams.Api.MessageExtensions.Response();
            }

            // Process settings data
            // Note: Settings property may not be available in current API
            log.Info("[MESSAGE_EXT_SETTING] Settings processing completed");

            return new Microsoft.Teams.Api.MessageExtensions.Response();
        }

        // Helper methods for creating responses
        private static Microsoft.Teams.Api.MessageExtensions.Response CreateSearchResults(string query, Microsoft.Teams.Common.Logging.ILogger log)
        {
            var attachments = new List<Microsoft.Teams.Api.MessageExtensions.Attachment>();

            // Create simple search results
            for (int i = 1; i <= 5; i++)
            {
                var card = new Microsoft.Teams.Cards.AdaptiveCard
                {
                    Body = new List<CardElement>
                    {
                        new TextBlock($"Search Result {i}")
                        {
                            Weight = TextWeight.Bolder,
                            Size = TextSize.Large
                        },
                        new TextBlock($"Query: '{query}' - Result description for item {i}")
                        {
                            Wrap = true,
                            IsSubtle = true
                        }
                    }
                };

                var previewCard = new ThumbnailCard()
                {
                    Title = $"Result {i}",
                    Text = $"This is a preview of result {i} for query '{query}'."
                };

                var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
                {
                    ContentType = Microsoft.Teams.Api.ContentType.AdaptiveCard,
                    Content = card,
                    Preview = new Microsoft.Teams.Api.MessageExtensions.Attachment
                    {
                        ContentType = Microsoft.Teams.Api.ContentType.ThumbnailCard,
                        Content = previewCard
                    }
                };

                attachments.Add(attachment);
            }

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Result,
                    AttachmentLayout = Microsoft.Teams.Api.Attachment.Layout.List,
                    Attachments = attachments
                }
            };
        }

        private static Microsoft.Teams.Api.MessageExtensions.Response HandleCreateCard(JsonElement? data, Microsoft.Teams.Common.Logging.ILogger log)
        {
            var title = GetJsonValue(data, "title") ?? "Default Title";
            var description = GetJsonValue(data, "description") ?? "Default Description";

            log.Info($"[CREATE_CARD] Title: {title}, Description: {description}");

            var card = new Microsoft.Teams.Cards.AdaptiveCard
            {
                Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                Body = new List<CardElement>
                {
                    new TextBlock("Custom Card Created")
                    {
                        Weight = TextWeight.Bolder,
                        Size = TextSize.Large,
                        Color = TextColor.Good
                    },
                    new TextBlock(title)
                    {
                        Weight = TextWeight.Bolder,
                        Size = TextSize.Medium
                    },
                    new TextBlock(description)
                    {
                        Wrap = true,
                        IsSubtle = true
                    }
                }
            };

            var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
            {
                ContentType = Microsoft.Teams.Api.ContentType.AdaptiveCard,
                Content = card
            };

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Result,
                    AttachmentLayout = Microsoft.Teams.Api.Attachment.Layout.List,
                    Attachments = new List<Microsoft.Teams.Api.MessageExtensions.Attachment> { attachment }
                }
            };
        }

        private static Microsoft.Teams.Api.MessageExtensions.Response HandleGetMessageDetails(Microsoft.Teams.Api.Activities.Invokes.MessageExtensions.SubmitActionActivity activity, Microsoft.Teams.Common.Logging.ILogger log)
        {
            var messageText = activity.Value?.MessagePayload?.Body?.Content ?? "No message content";
            var messageId = activity.Value?.MessagePayload?.Id ?? "Unknown";

            log.Info($"[GET_MESSAGE_DETAILS] Message ID: {messageId}");

            var card = new Microsoft.Teams.Cards.AdaptiveCard
            {
                Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                Body = new List<CardElement>
                {
                    new TextBlock("Message Details")
                    {
                        Weight = TextWeight.Bolder,
                        Size = TextSize.Large,
                        Color = TextColor.Accent
                    },
                    new TextBlock($"Message ID: {messageId}")
                    {
                        Wrap = true
                    },
                    new TextBlock($"Content: {messageText}")
                    {
                        Wrap = true
                    }
                }
            };

            var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
            {
                ContentType = new Microsoft.Teams.Api.ContentType("application/vnd.microsoft.card.adaptive"),
                Content = card
            };

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Result,
                    AttachmentLayout = Microsoft.Teams.Api.Attachment.Layout.List,
                    Attachments = new List<Microsoft.Teams.Api.MessageExtensions.Attachment> { attachment }
                }
            };
        }

        private static Microsoft.Teams.Api.MessageExtensions.Response CreateLinkUnfurlResponse(string url, Microsoft.Teams.Common.Logging.ILogger log)
        {
            var card = new Microsoft.Teams.Cards.AdaptiveCard
            {
                Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                Body = new List<CardElement>
                {
                    new TextBlock("Link Preview")
                    {
                        Weight = TextWeight.Bolder,
                        Size = TextSize.Medium
                    },
                    new TextBlock($"URL: {url}")
                    {
                        IsSubtle = true,
                        Wrap = true
                    },
                    new TextBlock("This is a preview of the linked content generated by the message extension.")
                    {
                        Wrap = true,
                        Size = TextSize.Small
                    }
                }
            };

            var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
            {
                ContentType = new Microsoft.Teams.Api.ContentType("application/vnd.microsoft.card.adaptive"),
                Content = card
            };

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Result,
                    AttachmentLayout = Microsoft.Teams.Api.Attachment.Layout.List,
                    Attachments = new List<Microsoft.Teams.Api.MessageExtensions.Attachment> { attachment }
                }
            };
        }

        private static Microsoft.Teams.Api.MessageExtensions.Response CreateItemSelectionResponse(object? selectedItem, Microsoft.Teams.Common.Logging.ILogger log)
        {
            var itemJson = JsonSerializer.Serialize(selectedItem);

            var card = new Microsoft.Teams.Cards.AdaptiveCard
            {
                Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                Body = new List<CardElement>
                {
                    new TextBlock("Item Selected")
                    {
                        Weight = TextWeight.Bolder,
                        Size = TextSize.Large,
                        Color = TextColor.Good
                    },
                    new TextBlock("You selected the following item:")
                    {
                        Wrap = true
                    },
                    new TextBlock(itemJson)
                    {
                        Wrap = true,
                        FontType = FontType.Monospace,
                        Separator = true
                    }
                }
            };

            var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
            {
                ContentType = new Microsoft.Teams.Api.ContentType("application/vnd.microsoft.card.adaptive"),
                Content = card
            };

            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Result,
                    AttachmentLayout = Microsoft.Teams.Api.Attachment.Layout.List,
                    Attachments = new List<Microsoft.Teams.Api.MessageExtensions.Attachment> { attachment }
                }
            };
        }

        private static Microsoft.Teams.Api.MessageExtensions.Response CreateErrorResponse(string message)
        {
            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Message,
                    Text = message
                }
            };
        }

        private static Microsoft.Teams.Api.MessageExtensions.Response CreateErrorActionResponse(string message)
        {
            return new Microsoft.Teams.Api.MessageExtensions.Response
            {
                ComposeExtension = new Microsoft.Teams.Api.MessageExtensions.Result
                {
                    Type = Microsoft.Teams.Api.MessageExtensions.ResultType.Message,
                    Text = message
                }
            };
        }

        private static string? GetJsonValue(JsonElement? data, string key)
        {
            if (data?.ValueKind == JsonValueKind.Object && data.Value.TryGetProperty(key, out var value))
            {
                return value.GetString();
            }
            return null;
        }

        private static Microsoft.Teams.Api.MessageExtensions.ActionResponse CreateFetchTaskResponse(string? commandId, Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info($"[CREATE_FETCH_TASK] Creating task for command: {commandId}");
            // Updated to use actual converation members

            // Create an adaptive card for the task module
            var card = new Microsoft.Teams.Cards.AdaptiveCard
            {
                Body = new List<CardElement>
                {
                    new TextBlock("Conversation Members is not implemented in C# yet :(")
                    {
                        Weight = TextWeight.Bolder,
                        Color = TextColor.Accent
                    },
                }
            };

            return new Microsoft.Teams.Api.MessageExtensions.ActionResponse
            {
                Task = new Microsoft.Teams.Api.TaskModules.ContinueTask(new Microsoft.Teams.Api.TaskModules.TaskInfo
                {
                    Title = "Fetch Task Dialog",
                    Height = new Microsoft.Teams.Common.Union<int, Microsoft.Teams.Api.TaskModules.Size>(Microsoft.Teams.Api.TaskModules.Size.Small),
                    Width = new Microsoft.Teams.Common.Union<int, Microsoft.Teams.Api.TaskModules.Size>(Microsoft.Teams.Api.TaskModules.Size.Small),
                    Card = new Microsoft.Teams.Api.Attachment(card)
                })
            };
        }

        private static string SanitizeForLog(string? input)
        {
            if (input == null) return "";
            return input.Replace("\r", "").Replace("\n", "");
        }
    }

    private static string GetSettingsHtml()
    {
        return """
<!DOCTYPE html>
<html>
<head>
    <title>Message Extension Settings</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://statics.teams.cdn.office.net/sdk/v1.12.0/js/MicrosoftTeams.min.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            max-width: 500px;
        }
        .form-group {
            margin-bottom: 15px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
        }
        select, input {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
        }
        .buttons {
            margin-top: 20px;
            text-align: right;
        }
        button {
            padding: 8px 16px;
            margin-left: 8px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        .btn-primary {
            background-color: #0078d4;
            color: white;
        }
        .btn-secondary {
            background-color: #6c757d;
            color: white;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>Message Extension Settings</h2>
        <form id="settingsForm">
            <div class="form-group">
                <label for="defaultAction">Default Action:</label>
                <select id="defaultAction" name="defaultAction">
                    <option value="search">Search</option>
                    <option value="compose">Compose</option>
                    <option value="both">Both</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="maxResults">Max Search Results:</label>
                <input type="number" id="maxResults" name="maxResults" value="10" min="1" max="50">
            </div>
            
            <div class="buttons">
                <button type="button" class="btn-secondary" onclick="cancelSettings()">Cancel</button>
                <button type="button" class="btn-primary" onclick="saveSettings()">Save</button>
            </div>
        </form>
    </div>

    <script>
        microsoftTeams.initialize();
        
        function saveSettings() {
            const formData = new FormData(document.getElementById('settingsForm'));
            const settings = {};
            for (let [key, value] of formData.entries()) {
                settings[key] = value;
            }
            
            microsoftTeams.tasks.submitTask(settings);
        }
        
        function cancelSettings() {
            microsoftTeams.tasks.submitTask();
        }
    </script>
</body>
</html>
""";
    }
}
