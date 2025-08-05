// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Defines a contract that represents a message response.
    /// E.g., returned by <see cref="INotificationTarget.SendAdaptiveCard"/> or <see cref="INotificationTarget.SendMessage"/>.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
    public class MessageResponse
    {
        /// <summary>
        /// The message ID.
        /// </summary>
        public string Id { get; set; }
    }
}
