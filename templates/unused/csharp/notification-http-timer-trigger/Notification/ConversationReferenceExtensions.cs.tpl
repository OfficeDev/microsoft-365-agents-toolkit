namespace {{SafeProjectName}}.Notification
{
    using Microsoft.Agents.Core.Models;
    using System.Text.Json;
    static internal class ConversationReferenceExtensions
    {
        static internal ConversationReference Clone(this ConversationReference reference)
        {
            if (reference == null)
            {
                return null;
            }

            return JsonSerializer.Deserialize<ConversationReference>(JsonSerializer.Serialize(reference));
        }

        static internal string GetKey(this ConversationReference reference)
        {
            return $"_{reference.Conversation?.TenantId}_{reference.Conversation?.Id}";
        }

        static internal string GetTargetType(this ConversationReference reference)
        {
            return reference?.Conversation?.ConversationType;
        }
    }
}
