// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Hosting.AspNetCore;

    /// <summary>
    /// Represents a bot to handle Adaptive Card Action Execute invoke activities.
    /// </summary>
    public class CardActionBot
    {
        private readonly CloudAdapter _adapter;
        private readonly CardActionMiddleware _middleware;

        /// <summary>
        /// Gets the registered card action handlers of this bot.
        /// </summary>
        public IList<IAdaptiveCardActionHandler> CardActionHandlers => _middleware.CardActionHandlers;

        /// <summary>
        /// Initializes a new instance of the <see cref="CardActionBot"/> class.
        /// </summary>
        /// <exception>
        /// <param name="adapter">The bot adapter.</param>
        /// <param name="options">The initialize options.</param>
        /// <paramref name="adapter"/> or <paramref name="options"/> is null.
        /// </exception>
        public CardActionBot(CloudAdapter adapter, CardActionOptions options)
        {
            _adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));

            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            _middleware = new CardActionMiddleware(options.Actions);
            _adapter.Use(_middleware);
        }

        /// <summary>
        /// Registers an adaptive card action handler to the conversation bot.
        /// </summary>
        /// <param name="cardActionHandler">A card action handler implements <see cref="IAdaptiveCardActionHandler"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="cardActionHandler"/> is null.</exception>
        public void RegisterHandler(IAdaptiveCardActionHandler cardActionHandler)
        {
            if (cardActionHandler == null)
            {
                throw new ArgumentNullException(nameof(cardActionHandler));
            }

            _middleware.CardActionHandlers.Add(cardActionHandler);
        }

        /// <summary>
        /// Registers a set of adaptive card action handlers to the conversation bot.
        /// </summary>
        /// <param name="cardActionHandlers">A list of adaptive card action handlers to  be registered to the bot.</param>
        /// <exception cref="ArgumentException"><paramref name="cardActionHandlers"/> is null or empty. </exception>
        public void RegisterHandlers(IList<IAdaptiveCardActionHandler> cardActionHandlers)
        {
            if (cardActionHandlers == null || !cardActionHandlers.Any())
            {
                throw new ArgumentException("There is no card action handler to be registered.", nameof(cardActionHandlers));
            }

            foreach (var handler in cardActionHandlers)
            {
                _middleware.CardActionHandlers.Add(handler);
            }
        }
    }
}
