// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;
    using Microsoft.Agents.Core.Models;

    /// <summary>
    /// Represents an adaptive card action handler to respond to an adaptiveCard/action invoke activity.
    /// </summary>
    public interface IAdaptiveCardActionHandler
    {
        /// <summary>
        /// The verb defined in adaptive card action that can trigger this handler.
        /// The verb string here is case-insensitive.
        /// </summary>
        string TriggerVerb { get; }

        /// <summary>
        /// Indicates the behavior for how the card response will be sent in Teams conversation.
        /// The default value is `AdaptiveCardResponse.ReplaceForInteractor`, which means the card
        /// response will replace the current one only for the interactor.
        /// </summary>
        AdaptiveCardResponse AdaptiveCardResponse { get; }

        /// <summary>
        /// The handler function that will be invoked when the card action is executed.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cardData">The card data object associated with the action.</param>
        /// <param name="cancellationToken"></param>
        /// <returns>Async task with invoke response.</returns>
        Task<InvokeResponse> HandleActionInvokedAsync(ITurnContext turnContext, object cardData, CancellationToken cancellationToken = default);
    }
}
