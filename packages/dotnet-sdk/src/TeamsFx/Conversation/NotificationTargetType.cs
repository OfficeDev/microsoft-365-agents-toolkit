// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// The target type where the notification will be sent to.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
    public enum NotificationTargetType
    {
        /// <summary>
        /// Represent an unknown target.
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// Represent a team channel. (By default, notification to a team will be sent to its "General" channel.)
        /// </summary>
        Channel,

        /// <summary>
        /// Represent a group chat.
        /// </summary>
        Group,

        /// <summary>
        /// Represent a personal chat.
        /// </summary>
        Person,
    }
}
