// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;

    /// <summary>
    /// Represents a simple text message for a command response.
    /// </summary>
    public class TextCommandResponse : ICommandResponse
    {
        /// <summary>
        /// Gets or sets the text messages used to reply to a command.
        /// </summary>
        /// <value>
        /// The text message.
        /// </value>
        public string Text { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="TextCommandResponse"/> class.
        /// </summary>
        /// <param name="text">The text of the message to send.</param>
        /// <exception cref="ArgumentNullException"><paramref name="text"/> is null or empty.</exception>
        public TextCommandResponse(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentNullException(nameof(text));
            }

            Text = text;
        }

        /// <inheritdoc/>
        public async Task SendResponseAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            await turnContext.SendActivityAsync(Text, cancellationToken: cancellationToken).ConfigureAwait(false);
        }
    }
}
