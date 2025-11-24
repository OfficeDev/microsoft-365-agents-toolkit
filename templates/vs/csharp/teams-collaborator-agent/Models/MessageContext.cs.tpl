using {{SafeProjectName}}.Storage;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Apps;

namespace {{SafeProjectName}}.Models
{
    public class ConversationMember
    {
        public string Name { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
    }

    public class CitationAppearance
    {
        public string Title { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public int AppearanceIndex { get; set; }
    }

    public class MessageContext
    {
        public string Text { get; set; } = string.Empty;
        public string ConversationId { get; set; } = string.Empty;
        public string? UserId { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string Timestamp { get; set; } = string.Empty;
        public bool IsPersonalChat { get; set; }
        public string ActivityId { get; set; } = string.Empty;
        public List<ConversationMember> Members { get; set; } = new();
        public ConversationMemory Memory { get; set; } = null!;
        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public List<CitationAppearance> Citations { get; set; } = new();
        public IContext<MessageActivity> TeamsContext { get; set; } = null!;
    }

    public static class MessageContextFactory
    {
        public static async Task<MessageContext> CreateAsync(
            IContext<MessageActivity> teamsContext,
            IDatabase database,
            ILogger? logger = null,
            bool includeMembers = false
        )
        {
            var activity = teamsContext.Activity;

            var text = activity.Text ?? string.Empty;
            var conversationId = activity.Conversation.Id;
            var userId = activity.From?.Id;
            var userName = activity.From?.Name ?? "User";
            var timestamp = activity.Timestamp?.ToString("o") ?? DateTime.UtcNow.ToString("o");
            var isPersonalChat = activity.Conversation.IsGroup != true;
            var activityId = activity.Id ?? string.Empty;

            var members = includeMembers
                ? await GetConversationMembersAsync(teamsContext, logger)
                : new List<ConversationMember>();

            var memory = new ConversationMemory(database, conversationId);

            var now = DateTime.UtcNow;
            var startTime = now.AddDays(-1).ToString("o");
            var endTime = now.ToString("o");

            return new MessageContext
            {
                Text = text,
                ConversationId = conversationId,
                UserId = userId,
                UserName = userName,
                Timestamp = timestamp,
                IsPersonalChat = isPersonalChat,
                ActivityId = activityId,
                Members = members,
                Memory = memory,
                StartTime = startTime,
                EndTime = endTime,
                Citations = new List<CitationAppearance>(),
                TeamsContext = teamsContext,
            };
        }

        public static async Task<List<ConversationMember>> GetConversationMembersAsync(
            IContext<MessageActivity> context,
            ILogger? logger = null
        )
        {
            try
            {
                if (context.Activity.Conversation.IsGroup != true)
                {
                    return new List<ConversationMember>
                    {
                        new ConversationMember
                        {
                            Name = context.Activity.From?.Name ?? "User",
                            Id = context.Activity.From?.Id ?? string.Empty,
                        },
                    };
                }

                var members = new List<ConversationMember>();

                if (context.Activity.From != null)
                {
                    members.Add(
                        new ConversationMember
                        {
                            Name = context.Activity.From.Name ?? "User",
                            Id = context.Activity.From.Id ?? string.Empty,
                        }
                    );
                }

                if (context.Activity.Recipient != null)
                {
                    members.Add(
                        new ConversationMember
                        {
                            Name = context.Activity.Recipient.Name ?? "Bot",
                            Id = context.Activity.Recipient.Id ?? string.Empty,
                        }
                    );
                }

                logger?.LogDebug(
                    "Retrieved {Count} members for conversation {ConversationId}",
                    members.Count,
                    context.Activity.Conversation.Id
                );

                return members;
            }
            catch (Exception ex)
            {
                logger?.LogWarning(ex, "Failed to retrieve conversation members");
                return new List<ConversationMember>();
            }
        }
    }
}
