// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Hosting.AspNetCore;

    /// <summary>
    /// Represents the options used to initialize a <see cref="ConversationBot"/>.
    /// </summary>
    public class ConversationOptions
    {
        /// <summary>
        /// Gets or sets the bot adapter 
        /// </summary>
        public CloudAdapter Adapter { get; set; }

        /// <summary>
        /// Gets or sets the notification option used to initialize <see cref="ConversationBot.Notification"/>.
        /// </summary>
        public NotificationOptions Notification { get; set; }

        /// <summary>
        /// Gets or sets the command option used to initialize the <see cref="ConversationBot.Command"/> .
        /// </summary>
        public CommandOptions Command { get; set; }

        /// <summary>
        /// Gets or sets the adaptive card action options used to initialize the <see cref="ConversationBot.CardAction"/> .
        /// </summary>
        public CardActionOptions CardAction { get; set; }
    }
}
