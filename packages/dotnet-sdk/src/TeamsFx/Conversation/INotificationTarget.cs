// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Represent a notification target.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Hosting.AspNetCore) instead.")]
    public interface INotificationTarget
    {
        /// <summary>
        /// The type of target.
        /// </summary>
        NotificationTargetType Type { get; }

        /// <summary>
        /// Send a plain text message.
        /// </summary>
        /// <param name="message">The plain text message.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The response of sending message.</returns>
        Task<MessageResponse> SendMessage(string message, CancellationToken cancellationToken = default);

        /// <summary>
        /// Send an adaptive card message.
        /// </summary>
        /// <param name="card">The adaptive card object.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The response of sending adaptive card message.</returns>
        Task<MessageResponse> SendAdaptiveCard(object card, CancellationToken cancellationToken = default);
    }
}
