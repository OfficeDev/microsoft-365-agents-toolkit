// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;

    /// <summary>
    /// Defines a contract that represents a command response.
    /// </summary>
    public interface ICommandResponse
    {
        /// <summary>
        /// Send the command response to the client.
        /// </summary>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        Task SendResponseAsync(ITurnContext turnContext, CancellationToken cancellationToken = default);
    }
}
