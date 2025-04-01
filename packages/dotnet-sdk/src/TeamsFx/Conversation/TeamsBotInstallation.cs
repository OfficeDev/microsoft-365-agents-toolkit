// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Connector;
    using Microsoft.Agents.Extensions.Teams.Models;
    using Microsoft.Agents.Hosting.AspNetCore;

    /// <summary>
    /// <para>
    /// An <see cref="INotificationTarget"/> that represents a bot installation. Teams Bot could be installed into:
    /// </para>
    /// <list type="bullet">
    ///     <item>
    ///         <description>Personal chat.</description>
    ///     </item>
    ///     <item>
    ///         <description>Group chat.</description>
    ///     </item>
    ///     <item>
    ///         <description>Team (by default the "General" channel).</description>
    ///     </item>
    /// </list>
    /// </summary>
    /// <remarks>
    /// It's recommended to get bot installations from <c>ConversationBot.Notification.GetInstallationsAsync</c>.
    /// </remarks>
    public class TeamsBotInstallation : INotificationTarget
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="botAppId">The application ID of the bot.</param>
        /// <param name="adapter">The bot adapter.</param>
        /// <param name="conversationReference">The <see cref="ConversationReference"/> of the bot installation.</param>
        /// <exception cref="ArgumentNullException">Throws if provided parameter is null.</exception>
        /// <remarks>
        /// It's recommended to get bot installations from <c>ConversationBot.Notification.GetInstallationsAsync</c>.
        /// </remarks>
        public TeamsBotInstallation(string botAppId, CloudAdapter adapter, ConversationReference conversationReference)
        {
            BotAppId = botAppId;
            Adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));
            ConversationReference = conversationReference ?? throw new ArgumentNullException(nameof(conversationReference));
            Type = ConversationReference.GetTargetType();
        }

        /// <summary>
        /// The bot adapter.
        /// </summary>
        public CloudAdapter Adapter { get; private set; }

        /// <summary>
        /// The application ID of the bot.
        /// </summary>
        public string BotAppId { get; private set; }

        /// <summary>
        /// The <see cref="ConversationReference"/> of the bot installation.
        /// </summary>
        public ConversationReference ConversationReference { get; private set; }

        /// <summary>
        /// The notification target type.
        /// <list type="bullet">
        ///     <item>
        ///         <description><see cref="NotificationTargetType.Channel"/> means bot is installed into a team and notification will be sent to its "General" channel.</description>
        ///     </item>
        ///     <item>
        ///         <description><see cref="NotificationTargetType.Group"/> means bot is installed into a group chat.</description>
        ///     </item>
        ///     <item>
        ///         <description><see cref="NotificationTargetType.Person"/> means bot is installed into a personal scope and notification will be sent to personal chat.</description>
        ///     </item>
        /// </list>
        /// </summary>
        public NotificationTargetType Type { get; private set; }

        /// <inheritdoc/>
        public async Task<MessageResponse> SendMessage(string message, CancellationToken cancellationToken = default)
        {
            var response = new MessageResponse();
            await Adapter.ContinueConversationAsync
            (
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var res = await context.SendActivityAsync(message, cancellationToken: ct).ConfigureAwait(false);
                    response.Id = res?.Id;
                },
                cancellationToken
            ).ConfigureAwait(false);
            return response;
        }

        /// <inheritdoc/>
        public async Task<MessageResponse> SendAdaptiveCard(object card, CancellationToken cancellationToken = default)
        {
            var response = new MessageResponse();
            await Adapter.ContinueConversationAsync
            (
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var res = await context.SendActivityAsync
                    (
                        MessageFactory.Attachment
                        (
                            new Attachment {
                                ContentType = "application/vnd.microsoft.card.adaptive",
                                Content = card,
                            }
                        ),
                        cancellationToken: ct
                    ).ConfigureAwait(false);
                    response.Id = res?.Id;
                },
                cancellationToken
            ).ConfigureAwait(false);
            return response;
        }

        /// <summary>
        /// Get channels from this bot installation.
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An array of channels if bot is installed into a team, otherwise returns an empty array.</returns>
        public async Task<Channel[]> GetChannelsAsync(CancellationToken cancellationToken = default)
        {
            var channels = new List<Channel>();
            if (Type != NotificationTargetType.Channel)
            {
                return channels.ToArray();
            }

            IList<ChannelInfo> teamsChannels = null;
            await Adapter.ContinueConversationAsync
            (
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var teamId = context.GetTeamsBotInstallationId();
                    if (teamId != null)
                    {
                        teamsChannels = await TeamsInfo.GetTeamChannelsAsync(context, teamId, ct).ConfigureAwait(false);
                    }
                },
                cancellationToken
            ).ConfigureAwait(false);

            if (teamsChannels != null)
            {
                foreach (var teamChannel in teamsChannels)
                {
                    channels.Add(new Channel(this, teamChannel));
                }
            }

            return channels.ToArray();
        }

        /// <summary>
        /// Get a pagined list of members from this bot installation.
        /// </summary>
        /// <param name="pageSize">Suggested number of entries on a page.</param>
        /// <param name="continuationToken">The continuation token.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An Array of members from where the bot is installed.</returns>
        public async Task<PagedData<Member>> GetPagedMembersAsync(
            int? pageSize = default,
            string continuationToken = default,
            CancellationToken cancellationToken = default)
        {
            PagedData<Member> result = null;
            await Adapter.ContinueConversationAsync(
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var pagedMembers = await TeamsInfo.GetPagedMembersAsync(context, pageSize, continuationToken, ct).ConfigureAwait(false);
                    result = new PagedData<Member> {
                        Data = pagedMembers.Members.Select(member => new Member(this, member)).ToArray(),
                        ContinuationToken = pagedMembers.ContinuationToken,
                    };
                },
                cancellationToken
            ).ConfigureAwait(false);

            return result;
        }

        /// <summary>
        /// Get members from this bot installation.
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An array of members from where the bot is installed.</returns>
        [Obsolete($"Use {nameof(GetPagedMembersAsync)} instead.")]
        public async Task<Member[]> GetMembersAsync(CancellationToken cancellationToken = default)
        {
            var members = new List<Member>();
            string continuationToken = null;
            do
            {
                var pagedData = await GetPagedMembersAsync(null, continuationToken, cancellationToken).ConfigureAwait(false);
                continuationToken = pagedData.ContinuationToken;
                members.AddRange(pagedData.Data);
            } while (!string.IsNullOrEmpty(continuationToken));

            return members.ToArray();
        }

        /// <summary>
        /// Get team details from this bot installation
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The team details if bot is installed into a team, otherwise returns null.</returns>
        public async Task<TeamDetails> GetTeamDetailsAsync(CancellationToken cancellationToken = default)
        {
            if (Type != NotificationTargetType.Channel)
            {
                return null;
            }

            TeamDetails teamDetails = null;
            await Adapter.ContinueConversationAsync
            (
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var teamId = context.GetTeamsBotInstallationId();
                    if (teamId != null)
                    {
                        teamDetails = await TeamsInfo.GetTeamDetailsAsync(context, teamId, ct).ConfigureAwait(false);
                    }
                },
                cancellationToken
            ).ConfigureAwait(false);

            return teamDetails;
        }
    }
}
