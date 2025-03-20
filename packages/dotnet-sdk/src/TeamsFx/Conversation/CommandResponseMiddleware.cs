// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;
    using Microsoft.Agents.Core.Models;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Middleware to handle message activity from Teams.
    /// </summary>
    public class CommandResponseMiddleware : IMiddleware
    {
        /// <summary>
        /// Gets or sets the list of <see cref="ITeamsCommandHandler"/> instances registered with the middleware.
        /// </summary>
        /// <value>
        /// The list of <see cref="ITeamsCommandHandler"/> instances registered with the middleware.
        /// </value>
        public IList<ITeamsCommandHandler> CommandHandlers { get; private set; } = new List<ITeamsCommandHandler>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandResponseMiddleware"/> class.
        /// </summary>
        /// <param name="commandHandlers">A list of command handlers.</param>
        public CommandResponseMiddleware(IList<ITeamsCommandHandler> commandHandlers = null)
        {
            if (commandHandlers != null && commandHandlers.Any())
            {
                CommandHandlers = commandHandlers;
            }
        }

        /// <inheritdoc/>
        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
        {
            var activityType = turnContext.Activity.Type;

            if (activityType == ActivityTypes.Message)
            {
                var receivedText = GetActivityText(turnContext.Activity);
                CommandMessage commandMessaage = new CommandMessage()
                {
                    Text = receivedText
                };

                foreach (var command in CommandHandlers)
                {
                    (bool shouldTrigger, commandMessaage.Matches) = ShouldTrigger(receivedText, command.TriggerPatterns);

                    if (shouldTrigger)
                    {
                        var response = await command.HandleCommandAsync(turnContext, commandMessaage, cancellationToken).ConfigureAwait(false);
                        if (response != null)
                        {
                            await response.SendResponseAsync(turnContext, cancellationToken).ConfigureAwait(false);
                        }

                        break;
                    }
                }
            }

            await next(cancellationToken).ConfigureAwait(false);
        }

        private static string GetActivityText(IActivity activity)
        {
            var text = activity.Text;
            var removedMentionText = activity.RemoveRecipientMention();
            if (!string.IsNullOrEmpty(removedMentionText))
            {
                text = activity.Text.Trim().ToLower();
            }

            return text;
        }

        private static (bool, MatchCollection) ShouldTrigger(string input, IEnumerable<ITriggerPattern> triggerPatterns)
        {
            var shouldTrigger = false;
            MatchCollection matches = null;

            foreach (var triggerPattern in triggerPatterns)
            {
                if (triggerPattern.ShouldTrigger(input))
                {
                    shouldTrigger = true;
                    if (triggerPattern is RegExpTrigger regexTrigger)
                    {
                        matches = regexTrigger.Pattern.Matches(input);
                    }

                    break;
                }
            }

            return (shouldTrigger, matches);
        }
    }
}
