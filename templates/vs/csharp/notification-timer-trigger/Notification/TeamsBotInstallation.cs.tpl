using Microsoft.Agents.Extensions.Teams.Connector;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Extensions.Teams.Models;
using Microsoft.Agents.Hosting.AspNetCore;

namespace {{SafeProjectName}}.Notification
{
    public class TeamsBotInstallation
    {
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

        public string Type { get; private set; }

        /// <inheritdoc/>
        public async Task<string> SendMessage(string message, CancellationToken cancellationToken = default)
        {
            string response = "";
            await Adapter.ContinueConversationAsync
            (
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var res = await context.SendActivityAsync(message, cancellationToken: ct).ConfigureAwait(false);
                    response = res?.Id;
                },
                cancellationToken
            ).ConfigureAwait(false);
            return response;
        }

        /// <inheritdoc/>
        public async Task<string> SendAdaptiveCard(object card, CancellationToken cancellationToken = default)
        {
            string response = "";
            await Adapter.ContinueConversationAsync
            (
                BotAppId,
                ConversationReference,
                async (context, ct) => {
                    var res = await context.SendActivityAsync
                    (
                        MessageFactory.Attachment
                        (
                            new Attachment
                            {
                                ContentType = "application/vnd.microsoft.card.adaptive",
                                Content = card,
                            }
                        ),
                        cancellationToken: ct
                    ).ConfigureAwait(false);
                    response = res?.Id;
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
            if (Type != "channel")
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
                    result = new PagedData<Member>
                    {
                        Data = pagedMembers.Members.Select(member => new Member(this, member)).ToArray(),
                        ContinuationToken = pagedMembers.ContinuationToken,
                    };
                },
                cancellationToken
            ).ConfigureAwait(false);

            return result;
        }

        /// <summary>
        /// Get team details from this bot installation
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The team details if bot is installed into a team, otherwise returns null.</returns>
        public async Task<TeamDetails> GetTeamDetailsAsync(CancellationToken cancellationToken = default)
        {
            if (Type != "channel")
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
