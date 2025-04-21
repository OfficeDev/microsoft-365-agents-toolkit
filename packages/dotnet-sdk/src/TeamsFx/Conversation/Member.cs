namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Connector;
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Models;

    /// <summary>
    /// An <see cref="INotificationTarget"/> that represents a team member.
    /// </summary>
    /// <remarks>
    /// It's recommended to get members from <see cref="TeamsBotInstallation.GetMembersAsync"/>.
    /// </remarks>
    public class Member : INotificationTarget
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="parent">The parent <see cref="TeamsBotInstallation"/> where this member is created from.</param>
        /// <param name="account">Detailed member account information.</param>
        /// <exception cref="ArgumentNullException">Throws if provided parameter is null.</exception>
        /// <remarks>
        /// It's recommended to get members from <see cref="TeamsBotInstallation.GetMembersAsync"/>.
        /// </remarks>
        public Member(TeamsBotInstallation parent, TeamsChannelAccount account)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            Account = account ?? throw new ArgumentNullException(nameof(account));
        }

        /// <summary>
        /// The parent <see cref="TeamsBotInstallation"/> where this member is created from.
        /// </summary>
        public TeamsBotInstallation Parent { get; private set; }

        /// <summary>
        /// Detailed member account information.
        /// </summary>
        public TeamsChannelAccount Account { get; private set; }

        /// <summary>
        /// The type of target. For member it's always <see cref="NotificationTargetType.Person"/>.
        /// </summary>
        public NotificationTargetType Type { get => NotificationTargetType.Person; }

        /// <inheritdoc/>
        public async Task<MessageResponse> SendMessage(string message, CancellationToken cancellationToken = default)
        {
            var response = new MessageResponse();
            await Parent.Adapter.ContinueConversationAsync
            (
                Parent.BotAppId,
                Parent.ConversationReference,
                async (context1, ct1) => {
                    var conversation = await NewConversation(context1, ct1).ConfigureAwait(false);
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
                    var conversation = await NewConversation(context1, ct1).ConfigureAwait(false);
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
                            cancellationToken: ct2).ConfigureAwait(false);
                            response.Id = res?.Id;
                        },
                        ct1
                    ).ConfigureAwait(false);
                },
                cancellationToken
            ).ConfigureAwait(false);
            return response;
        }

        private async Task<ConversationReference> NewConversation(ITurnContext context, CancellationToken cancellationToken = default)
        {
            var reference = context.Activity.GetConversationReference();
            var personConversation = reference.Clone();
            var connectorClient = context.Services.Get<IConnectorClient>();
            var conversation = await connectorClient.Conversations.CreateConversationAsync
            (
                new ConversationParameters {
                    IsGroup = false,
                    Agent = context.Activity.Recipient,
                    Members = new List<ChannelAccount>() { Account },
                    ChannelData = { },
                    TenantId = context.Activity.Conversation.TenantId,
                },
                cancellationToken
            ).ConfigureAwait(false);
            personConversation.Conversation.Id = conversation.Id;
            return personConversation;
        }
    }
}
