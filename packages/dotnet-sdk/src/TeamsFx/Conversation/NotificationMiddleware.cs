// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Models;

    internal class NotificationMiddleware : IMiddleware
    {
        private readonly IConversationReferenceStore _store;

        public NotificationMiddleware(IConversationReferenceStore store)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
        {
            var activityType = ClassifyActivity(turnContext.Activity);
            switch (activityType)
            {
                case ActivityType.CurrentBotInstalled:
                case ActivityType.TeamRestored:
                    {
                        var reference = turnContext.Activity.GetConversationReference();
                        var options = new ConversationReferenceStoreAddOptions {
                            Overwrite = true
                        };
                        await _store.Add(reference.GetKey(), reference, options, cancellationToken).ConfigureAwait(false);
                        break;
                    }
                case ActivityType.CurrentBotMessaged:
                    {
                        await TryAddMessagedReference(turnContext, cancellationToken).ConfigureAwait(false);
                        break;
                    }
                case ActivityType.CurrentBotUninstalled:
                case ActivityType.TeamDeleted:
                    {
                        var reference = turnContext.Activity.GetConversationReference();
                        await _store.Remove(reference.GetKey(), reference, cancellationToken).ConfigureAwait(false);
                        break;
                    }
                default:
                    {
                        break;
                    }
            }

            await next(cancellationToken).ConfigureAwait(false);
        }

        private ActivityType ClassifyActivity(IActivity activity)
        {
            var activityType = activity?.Type;
            if (ActivityTypes.InstallationUpdate.Equals(activityType, StringComparison.OrdinalIgnoreCase))
            {
                var action = activity.Action;
                if ("add".Equals(action, StringComparison.OrdinalIgnoreCase))
                {
                    return ActivityType.CurrentBotInstalled;
                }
                else
                {
                    return ActivityType.CurrentBotUninstalled;
                }
            }
            else if (ActivityTypes.ConversationUpdate.Equals(activityType, StringComparison.OrdinalIgnoreCase))
            {
                var eventType = activity.GetChannelData<TeamsChannelData>()?.EventType;
                if ("teamDeleted".Equals(eventType, StringComparison.OrdinalIgnoreCase))
                {
                    return ActivityType.TeamDeleted;
                }
                else if ("teamRestored".Equals(eventType, StringComparison.OrdinalIgnoreCase))
                {
                    return ActivityType.TeamRestored;
                }
            }
            else if (ActivityTypes.Message.Equals(activityType, StringComparison.OrdinalIgnoreCase))
            {
                return ActivityType.CurrentBotMessaged;
            }

            return ActivityType.Unknown;
        }

        private async Task<bool> TryAddMessagedReference(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            var reference = turnContext.Activity.GetConversationReference();
            var conversationType = reference?.Conversation?.ConversationType;
            if ("personal".Equals(conversationType, StringComparison.OrdinalIgnoreCase) || "groupChat".Equals(conversationType, StringComparison.OrdinalIgnoreCase))
            {
                var options = new ConversationReferenceStoreAddOptions {
                    Overwrite = false
                };
                var isUpdated = await _store.Add(reference.GetKey(), reference, options, cancellationToken).ConfigureAwait(false);
                return isUpdated;
            }
            else if ("channel".Equals(conversationType, StringComparison.OrdinalIgnoreCase))
            {
                var channelData = turnContext.Activity.GetChannelData<TeamsChannelData>();
                var teamId = channelData?.Team?.Id;
                var channelId = channelData?.Channel?.Id;
                // `teamId == channelId` means General channel. Ignore messaging in non-General channel.
                if (teamId != null && (channelId == null || string.Equals(teamId, channelId)))
                {
                    var channelReference = reference.Clone();
                    channelReference.Conversation.Id = teamId;

                    var options = new ConversationReferenceStoreAddOptions {
                        Overwrite = false
                    };
                    var isUpdated = await _store.Add(channelReference.GetKey(), channelReference, options, cancellationToken).ConfigureAwait(false);
                    return isUpdated;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
    }
}
