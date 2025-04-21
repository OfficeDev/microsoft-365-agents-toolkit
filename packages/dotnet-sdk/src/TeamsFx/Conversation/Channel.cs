namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Models;

    /// <summary>
    /// An <see cref="INotificationTarget"/> that represents a team channel.
    /// </summary>
    /// <remarks>
    /// It's recommended to get channels from <see cref="TeamsBotInstallation.GetChannelsAsync"/>.
    /// </remarks>
    public class Channel : INotificationTarget
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="parent">The parent <see cref="TeamsBotInstallation"/> where this channel is created from.</param>
        /// <param name="info">Detailed channel information.</param>
        /// <exception cref="ArgumentNullException">Throws if provided parameter is null.</exception>
        /// <remarks>
        /// It's recommended to get channels from <see cref="TeamsBotInstallation.GetChannelsAsync"/>.
        /// </remarks>
        public Channel(TeamsBotInstallation parent, ChannelInfo info)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            Info = info ?? throw new ArgumentNullException(nameof(info));
        }

        /// <summary>
        /// The parent <see cref="TeamsBotInstallation"/> where this channel is created from.
        /// </summary>
        public TeamsBotInstallation Parent { get; private set; }

        /// <summary>
        /// Detailed channel information.
        /// </summary>
        public ChannelInfo Info { get; private set; }

        /// <summary>
        /// The type of target. For channel it's always <see cref="NotificationTargetType.Channel"/>.
        /// </summary>
        public NotificationTargetType Type { get => NotificationTargetType.Channel; }

        /// <inheritdoc/>
        public async Task<MessageResponse> SendMessage(string message, CancellationToken cancellationToken = default)
        {
            var response = new MessageResponse();
            await Parent.Adapter.ContinueConversationAsync
            (
                Parent.BotAppId,
                Parent.ConversationReference,
                async (context1, ct1) => {
                    var conversation = NewConversation(context1);
                    await Parent.Adapter.ContinueConversationAsync
                    (
                        Parent.BotAppId,
                        conversation,
                        async (context2, ct2) => {
                            var res = await context2.SendActivityAsync(message, cancellationToken: ct2).ConfigureAwait(false);
                            response.Id = res?.Id;
                        },
                        ct1
                    ).ConfigureAwait(false);
                },
                cancellationToken
            ).ConfigureAwait(false);
            return response;
        }

        /// <inheritdoc/>
        public async Task<MessageResponse> SendAdaptiveCard(object card, CancellationToken cancellationToken = default)
        {
            var response = new MessageResponse();
            await Parent.Adapter.ContinueConversationAsync
            (
                Parent.BotAppId,
                Parent.ConversationReference,
                async (context1, ct1) => {
                    var conversation = NewConversation(context1);
                    await Parent.Adapter.ContinueConversationAsync
                    (
                        Parent.BotAppId,
                        conversation,
                        async (context2, ct2) => {
                            var res = await context2.SendActivityAsync
                            (
                                MessageFactory.Attachment
                                (
                                    new Attachment {
                                        ContentType = "application/vnd.microsoft.card.adaptive",
                                        Content = card,
                                    }
                                ),
                                cancellationToken: ct2
                            ).ConfigureAwait(false);
                            response.Id = res?.Id;
                        },
                        ct1
                    ).ConfigureAwait(false);
                },
                cancellationToken
            ).ConfigureAwait(false);
            return response;
        }

        private ConversationReference NewConversation(ITurnContext context)
        {
            var reference = context.Activity.GetConversationReference();
            var channelConversation = reference.Clone();
            channelConversation.Conversation.Id = Info.Id ?? string.Empty;
            return channelConversation;
        }
    }
}
