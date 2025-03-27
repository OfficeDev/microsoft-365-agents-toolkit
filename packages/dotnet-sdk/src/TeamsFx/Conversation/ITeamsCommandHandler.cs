// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;

    /// <summary>
    /// Defines a contract that represents a command handler that can handle commands received from Teams.
    /// </summary>
    public interface ITeamsCommandHandler
    {

        /// <summary>
        /// Gets a set of string or regular expression pattern used to trigger the command handler.
        /// If not provided, the handler will be triggered by the command name.
        /// </summary>
        /// <remarks>
        /// The command handler can be triggered if
        /// <list type="bullet">
        ///     <item>
        ///         <description> the input command equals to the one of the string patterns.</description>
        ///     </item>
        ///     <item>
        ///         <description> the input command matches one of the regular expression patterns.</description>
        ///     </item>
        /// </list>
        /// </remarks>
        IEnumerable<ITriggerPattern> TriggerPatterns { get; }

        /// <summary>
        /// Handles the command.
        /// </summary>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="message">The command message.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns></returns>
        Task<ICommandResponse> HandleCommandAsync(ITurnContext turnContext, CommandMessage message, CancellationToken cancellationToken = default);
    }
}
