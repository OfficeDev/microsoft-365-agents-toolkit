// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;
    using Microsoft.Agents.Core.Models;

    /// <summary>
    /// Represents a message activity for a command response.
    /// </summary>
    public class ActivityCommandResponse : ICommandResponse
    {
        /// <summary>
        /// Gets or sets the activity used to reply to a command.
        /// </summary>
        /// <value>
        /// The activity used to reply to a command
        /// </value>
        public IActivity Activity { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ActivityCommandResponse"/> class.
        /// </summary>
        /// <param name="activity"></param>
        /// <exception cref="ArgumentNullException"><paramref name="activity"/> is null.</exception>
        public ActivityCommandResponse(IActivity activity)
        {
            Activity = activity ?? throw new ArgumentNullException(nameof(activity));
        }

        /// <inheritdoc/>
        public async Task SendResponseAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            await turnContext.SendActivityAsync(Activity, cancellationToken).ConfigureAwait(false);
        }
    }
}
