// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Options to initialize a <see cref="CommandBot"/>.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
    public class CardActionOptions
    {
        /// <summary>
        /// Gets or sets a set of adaptive card action handlers used to process card actions for this bot.
        /// </summary>
        /// <value>
        /// The card action handlers used to process command.
        /// </value>
        public IList<IAdaptiveCardActionHandler> Actions { get; set; }
    }
}
