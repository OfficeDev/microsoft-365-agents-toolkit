using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Storage;
using Microsoft.Teams.Api.Activities;

namespace {{SafeProjectName}}.Utils
{
    public static class MessageUtils
    {
        public static MessageActivity FinalizePromptResponse(
            string text,
            MessageContext context,
            ILogger logger
        )
        {
            var finalText = text;

            if (context.Citations != null && context.Citations.Count > 0)
            {
                logger.LogDebug(
                    "Adding {Count} citations to message activity",
                    context.Citations.Count
                );

                for (int i = 0; i < context.Citations.Count; i++)
                {
                    finalText += $" [{i + 1}]";
                }
            }

            var messageActivity = new MessageActivity(finalText);
            messageActivity.AddAIGenerated();
            messageActivity.AddFeedback();

            return messageActivity;
        }

        public static List<MessageRecord> CreateMessageRecords(
            IEnumerable<MessageActivity> activities
        )
        {
            var activitiesList = activities.ToList();
            if (activitiesList.Count == 0)
            {
                return new List<MessageRecord>();
            }

            var conversationId = activitiesList[0].Conversation.Id;

            return activitiesList
                .Select(activity =>
                {
                    var isAiGenerated =
                        activity.Entities?.Any(e =>
                            {
                                var typeStr = e.Type?.ToString() ?? string.Empty;
                                return typeStr.Contains("AIGenerated", StringComparison.OrdinalIgnoreCase)
                                    || typeStr.Contains("AIGeneratedContent", StringComparison.OrdinalIgnoreCase);
                            }) ?? false;

                    var role = isAiGenerated ? "assistant" : "user";
                    var content =
                        activity.Text?.Replace("<at>", string.Empty)
                            .Replace("</at>", string.Empty)
                            .Trim()
                        ?? string.Empty;

                    return new MessageRecord
                    {
                        ConversationId = conversationId,
                        Role = role,
                        Content = content,
                        Timestamp = activity.Timestamp?.ToString("o") ?? DateTime.UtcNow.ToString("o"),
                        ActivityId = activity.Id ?? string.Empty,
                        Name = activity.From?.Name ?? "Collaborator",
                        UserId = activity.From?.Id ?? string.Empty,
                    };
                })
                .ToList();
        }

        public static MessageRecord CreateMessageRecord(
            MessageActivity activity,
            string role = "user"
        )
        {
            var content =
                activity.Text?.Replace("<at>", string.Empty).Replace("</at>", string.Empty)
                ?? string.Empty;

            return new MessageRecord
            {
                ConversationId = activity.Conversation.Id,
                Role = role,
                Content = content,
                Timestamp = activity.Timestamp?.ToString("o") ?? DateTime.UtcNow.ToString("o"),
                ActivityId = activity.Id ?? string.Empty,
                Name = activity.From?.Name ?? "User",
                UserId = activity.From?.Id ?? string.Empty,
            };
        }
    }
}
