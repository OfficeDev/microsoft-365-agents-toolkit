// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Connector;
    using Microsoft.Agents.Extensions.Teams.Models;
    using Microsoft.Agents.Hosting.AspNetCore;
    using System.Net;

    /// <summary>
    /// Provide utilities to send notification to varies targets (e.g., member, group, channel).
    /// </summary>
    public class NotificationBot
    {
        private readonly CloudAdapter _adapter;
        private readonly string _botAppId;
        private readonly IConversationReferenceStore _store;

        /// <summary>
        /// Create new instance of the <see cref="NotificationBot"/>.
        /// </summary>
        /// <param name="adapter">The bot adapter.</param>
        /// <param name="options">The initialize options.</param>
        /// <exception cref="ArgumentNullException">Throws if provided parameter is null.</exception>
        /// <exception cref="ArgumentException">Throws if provided parameter is invalid.</exception>
        public NotificationBot(CloudAdapter adapter, NotificationOptions options)
        {
            _adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));

            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            _botAppId = options.BotAppId;

            if (options.Store != null)
            {
                _store = options.Store;
            }
#pragma warning disable CS0618, CS0612 // Suppress warning for obsolete type or member
            else if (options.Storage != null)
            {
                _store = new DefaultConversationReferenceStore(options.Storage);
            }
            else
            {
                INotificationTargetStorage storage;
                var onAzure = Environment.GetEnvironmentVariable("RUNNING_ON_AZURE");
                if ("1".Equals(onAzure))
                {
                    storage = new LocalFileStorage(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP") ?? Environment.CurrentDirectory));
                }
                else
                {
                    storage = new LocalFileStorage(Path.GetFullPath(Environment.GetEnvironmentVariable("TEAMSFX_NOTIFICATION_LOCALSTORE_DIR") ?? Environment.CurrentDirectory));
                }
                _store = new DefaultConversationReferenceStore(storage);
            }
#pragma warning restore CS0618, CS0612 // Restore warning
            _adapter.Use(new NotificationMiddleware(_store));
        }

        /// <summary>
        /// Create a <see cref="TeamsBotInstallation"/> instance with conversation reference.
        /// </summary>
        /// <param name="reference">The <see cref="ConversationReference"/> of the bot installation.</param>
        /// <returns>The <see cref="TeamsBotInstallation"/> instance.</returns>
        /// <exception cref="ArgumentNullException">Throws if provided parameter is null.</exception>
        public TeamsBotInstallation BuildTeamsBotInstallation(ConversationReference reference)
        {
            if (reference == null)
            {
                throw new ArgumentNullException(nameof(reference));
            }

            return new TeamsBotInstallation(_botAppId, _adapter, reference);
        }

        /// <summary>
        /// Validate the installation by getting paged members.
        /// </summary>
        /// <param name="reference">The <see cref="ConversationReference"/> of the bot installation.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>Returns false if receives "BotNotInConversationRoster" error, otherwise returns true</returns>
        public async Task<bool> ValidateInstallationAsync(ConversationReference reference, CancellationToken cancellationToken = default) 
        {
            var isValid = true;
            await _adapter.ContinueConversationAsync
            (
                _botAppId,
                reference,
                async (context, ct) => {
                    try
                    {
                        // try get member to see if the installation is still valid
                        await TeamsInfo.GetPagedMembersAsync(context, 1, null, ct).ConfigureAwait(false);
                    }
                    catch (Exception e)
                    {
                        if (e is OperationCanceledException exception)
                        {
                            var status = exception.HResult;
                            var error = exception.Message ?? string.Empty;
                            if (status == (int)HttpStatusCode.Forbidden && error.Contains("BotNotInConversationRoster"))
                            {
                                // bot is uninstalled
                                isValid = false;
                            }
                        }
                    }
                },
                cancellationToken
            ).ConfigureAwait(false);
            return isValid;
        }

        /// <summary>
        /// Get a paginated list of targets where the bot is installed.
        /// </summary>
        /// <param name="pageSize">Suggested number of entries on a page.</param>
        /// <param name="continuationToken">The continuation token.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <param name="validationEnabled">The parameter to enable or disable installation validation.</param>
        /// <returns>A paginated list of <see cref="TeamsBotInstallation"/>.</returns>
        /// <remarks>
        /// The result is retrieving from the persisted storage.
        /// </remarks>
        public async Task<PagedData<TeamsBotInstallation>> GetPagedInstallationsAsync(
            int? pageSize = default,
            string continuationToken = default,
            CancellationToken cancellationToken = default,
            bool validationEnabled = true)
        {
            var pagedData = await _store.List(pageSize, continuationToken, cancellationToken).ConfigureAwait(false);
            var installations = new List<TeamsBotInstallation>();

            foreach (var reference in pagedData.Data)
            {
                // validate connection
                bool valid = true;
                if (validationEnabled)
                {
                    valid = await ValidateInstallationAsync(reference, cancellationToken).ConfigureAwait(false);
                }
                if (!validationEnabled || (validationEnabled && valid))
                {
                    installations.Add(new TeamsBotInstallation(_botAppId, _adapter, reference));
                }
                else
                {
                    await _store.Remove(reference.GetKey(), reference, cancellationToken).ConfigureAwait(false);
                }
            }

            return new PagedData<TeamsBotInstallation> {
                Data = installations.ToArray(),
                ContinuationToken = pagedData.ContinuationToken,
            };
        }

        /// <summary>
        /// Get all targets where the bot is installed.
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An array of <see cref="TeamsBotInstallation"/>.</returns>
        /// <remarks>
        /// The result is retrieving from the persisted storage.
        /// </remarks>
        [Obsolete($"Use {nameof(GetPagedInstallationsAsync)} instead.")]
        public async Task<TeamsBotInstallation[]> GetInstallationsAsync(CancellationToken cancellationToken = default)
        {
            string continuationToken = null;
            var installations = new List<TeamsBotInstallation>();
            do
            {
                var pagedData = await GetPagedInstallationsAsync(null, continuationToken, cancellationToken).ConfigureAwait(false);
                installations.AddRange(pagedData.Data);
                continuationToken = pagedData.ContinuationToken;
            } while (!string.IsNullOrEmpty(continuationToken));

            return installations.ToArray();
        }

        /// <summary>
        /// Returns the first <see cref="Member"/> where predicate is true, and null otherwise.
        /// </summary>
        /// <param name="predicate">
        /// Find calls predicate once for each member of the installation, 
        /// until it finds one where predicate returns true. If such a member is found, 
        /// find immediately returns that member.Otherwise, find returns null.
        /// </param>
        /// <param name="scope">The scope to find members from the installations. 
        /// (personal chat, group chat, Teams channel)
        /// </param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The first <see cref="Member"/> where predicate is true, and null otherwise.</returns>
        /// <exception cref="ArgumentNullException">Throws when predicate is null.</exception>
        public async Task<Member> FindMemberAsync(
            Func<Member, Task<bool>> predicate,
            SearchScope scope = SearchScope.All,
            CancellationToken cancellationToken = default)
        {
            if (predicate == null)
            {
                throw new ArgumentNullException(nameof(predicate));
            }

            string installationContinuationToken = null;
            do
            {
                var pagedInstallations = await GetPagedInstallationsAsync(null, installationContinuationToken, cancellationToken).ConfigureAwait(false);
                installationContinuationToken = pagedInstallations.ContinuationToken;
                foreach (var target in pagedInstallations.Data)
                {
                    if (MatchSearchScope(target, scope))
                    {
                        string memberContinuationToken = null;
                        do
                        {
                            var pagedMembers = await target.GetPagedMembersAsync(null, memberContinuationToken, cancellationToken).ConfigureAwait(false);
                            memberContinuationToken = pagedMembers.ContinuationToken;
                            foreach (var member in pagedMembers.Data)
                            {
                                if (await predicate(member).ConfigureAwait(false))
                                {
                                    return member;
                                }
                            }
                        } while (!string.IsNullOrEmpty(memberContinuationToken));

                    }
                }
            } while (!string.IsNullOrEmpty(installationContinuationToken));


            return null;
        }

        /// <summary>
        /// Returns all <see cref="Member"/> where predicate is true, and empty array otherwise.
        /// </summary>
        /// <param name="predicate">Find calls predicate for each member of the installation.</param>
        /// <param name="scope">The scope to find members from the installations. 
        /// (personal chat, group chat, Teams channel)
        /// </param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An array of <see cref="Member"/> where predicate is true, and empty array otherwise.</returns>
        /// <exception cref="ArgumentNullException">Throws when predicate is null.</exception>
        public async Task<Member[]> FindAllMembersAsync(
            Func<Member, Task<bool>> predicate,
            SearchScope scope = SearchScope.All,
            CancellationToken cancellationToken = default)
        {
            if (predicate == null)
            {
                throw new ArgumentNullException(nameof(predicate));
            }

            var result = new List<Member>();
            string installationContinuationToken = null;
            do
            {
                var pagedInstallations = await GetPagedInstallationsAsync(null, installationContinuationToken, cancellationToken).ConfigureAwait(false);
                installationContinuationToken = pagedInstallations.ContinuationToken;
                foreach (var target in pagedInstallations.Data)
                {
                    if (MatchSearchScope(target, scope))
                    {
                        string memberContinuationToken = null;
                        do
                        {
                            var pagedMembers = await target.GetPagedMembersAsync(null, memberContinuationToken, cancellationToken).ConfigureAwait(false);
                            memberContinuationToken = pagedMembers.ContinuationToken;
                            foreach (var member in pagedMembers.Data)
                            {
                                if (await predicate(member).ConfigureAwait(false))
                                {
                                    result.Add(member);
                                }
                            }
                        } while (!string.IsNullOrEmpty(memberContinuationToken));
                    }
                }
            } while (!string.IsNullOrEmpty(installationContinuationToken));

            return result.ToArray();
        }

        /// <summary>
        /// Returns the first <see cref="Channel"/> where predicate is true, and null otherwise.
        /// (Ensure the bot app is installed into the `General` channel, otherwise null will be returned.)
        /// </summary>
        /// <param name="predicate">
        /// Find calls predicate once for each channel of the installation, 
        /// until it finds one where predicate returns true. If such a channel is found, 
        /// find immediately returns that channel.Otherwise, find returns null.
        /// </param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The first <see cref="Channel"/> where predicate is true, and null otherwise.</returns>
        /// <exception cref="ArgumentNullException">Throws when predicate is null.</exception>
        public async Task<Channel> FindChannelAsync(
            Func<Channel, TeamDetails, Task<bool>> predicate,
            CancellationToken cancellationToken = default)
        {
            if (predicate == null)
            {
                throw new ArgumentNullException(nameof(predicate));
            }

            string continuationToken = null;
            do
            {
                var pagedInstallations = await GetPagedInstallationsAsync(null, continuationToken, cancellationToken).ConfigureAwait(false);
                continuationToken = pagedInstallations.ContinuationToken;
                foreach (var target in pagedInstallations.Data)
                {
                    if (target.Type == NotificationTargetType.Channel)
                    {
                        var teamDetails = await target.GetTeamDetailsAsync(cancellationToken).ConfigureAwait(false);
                        var channels = await target.GetChannelsAsync(cancellationToken).ConfigureAwait(false);
                        foreach (var channel in channels)
                        {
                            if (await predicate(channel, teamDetails).ConfigureAwait(false))
                            {
                                return channel;
                            }
                        }
                    }
                }
            } while (!string.IsNullOrEmpty(continuationToken));

            return null;
        }

        /// <summary>
        /// Returns all <see cref="Channel"/> where predicate is true, and empty array otherwise.
        /// (Ensure the bot app is installed into the `General` channel, otherwise empty array will be returned.)
        /// </summary>
        /// <param name="predicate">Predicate find calls predicate for each channel of the installation.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An array of <see cref="Channel"/> where predicate is true, and empty array otherwise.</returns>
        /// <exception cref="ArgumentNullException">Throws when predicate is null.</exception>
        public async Task<Channel[]> FindAllChannelsAsync(
            Func<Channel, TeamDetails, Task<bool>> predicate,
            CancellationToken cancellationToken = default)
        {
            if (predicate == null)
            {
                throw new ArgumentNullException(nameof(predicate));
            }

            var result = new List<Channel>();
            string continuationToken = null;
            do
            {
                var pagedInstallations = await GetPagedInstallationsAsync(null, continuationToken, cancellationToken).ConfigureAwait(false);
                continuationToken = pagedInstallations.ContinuationToken;
                foreach (var target in pagedInstallations.Data)
                {
                    if (target.Type == NotificationTargetType.Channel)
                    {
                        var teamDetails = await target.GetTeamDetailsAsync(cancellationToken).ConfigureAwait(false);
                        var channels = await target.GetChannelsAsync(cancellationToken).ConfigureAwait(false);
                        foreach (var channel in channels)
                        {
                            if (await predicate(channel, teamDetails).ConfigureAwait(false))
                            {
                                result.Add(channel);
                            }
                        }
                    }
                }
            } while (!string.IsNullOrEmpty(continuationToken));

            return result.ToArray();
        }

        private static bool MatchSearchScope(TeamsBotInstallation target, SearchScope scope = SearchScope.All)
        {
            return target.Type switch {
                NotificationTargetType.Channel => scope.HasFlag(SearchScope.Channel),
                NotificationTargetType.Person => scope.HasFlag(SearchScope.Person),
                NotificationTargetType.Group => scope.HasFlag(SearchScope.Group),
                _ => false,
            };
        }
    }
}
