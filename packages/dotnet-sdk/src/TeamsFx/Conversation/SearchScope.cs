// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// The search scope when calling <see cref="NotificationBot.FindMemberAsync"/> and <see cref="NotificationBot.FindAllMembersAsync"/>.
    /// The search scope is a flagged enum and it can be combined with `|`.
    /// For example, to search from personal chat and group chat, use `SearchScope.Person | SearchScope.Group`.
    /// </summary>
    [Flags]
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
    public enum SearchScope
    {
        /// <summary>
        /// Search members from the installations in personal chat only.
        /// </summary>
        Person = 1,

        /// <summary>
        /// Search members from the installations in group chat only.
        /// </summary>
        Group = 2,

        /// <summary>
        /// Search members from the installations in Teams channel only.
        /// </summary>
        Channel = 4,

        /// <summary>
        /// Search members from all installations including personal chat, group chat and Teams channel.
        /// </summary>
        All = Person | Group | Channel
    }
}
