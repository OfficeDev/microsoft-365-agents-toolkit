// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
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
            return $"_{ reference.Conversation?.TenantId}_{ reference.Conversation?.Id}";
        }

        static internal NotificationTargetType GetTargetType(this ConversationReference reference)
        {
            var conversationType = reference?.Conversation?.ConversationType;
            if ("personal".Equals(conversationType, StringComparison.OrdinalIgnoreCase))
            {
                return NotificationTargetType.Person;
            }
            else if ("groupChat".Equals(conversationType, StringComparison.OrdinalIgnoreCase))
            {
                return NotificationTargetType.Group;
            }
            else if ("channel".Equals(conversationType, StringComparison.OrdinalIgnoreCase))
            {
                return NotificationTargetType.Channel;
            }
            else
            {
                return NotificationTargetType.Unknown;
            }
        }
    }
}
