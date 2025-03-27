// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.BotBuilder;
    using Microsoft.Agents.Core.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using System.Threading;
    using System.Threading.Tasks;

    internal class CardActionMiddleware : IMiddleware
    {
        internal readonly string defaultMessage = "Your response was sent to the app";

        /// <summary>
        /// Gets or sets the list of <see cref="IAdaptiveCardActionHandler"/> instances registered with the middleware.
        /// </summary>
        /// <value>
        /// The list of <see cref="IAdaptiveCardActionHandler"/> instances registered with the middleware.
        /// </value>
        public IList<IAdaptiveCardActionHandler> CardActionHandlers { get; private set; } = new List<IAdaptiveCardActionHandler>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandResponseMiddleware"/> class.
        /// </summary>
        /// <param name="cardActionHandlers">A list of command handlers.</param>
        public CardActionMiddleware(IList<IAdaptiveCardActionHandler> cardActionHandlers = null)
        {
            if (cardActionHandlers != null && cardActionHandlers.Any())
            {
                CardActionHandlers = cardActionHandlers;
            }
        }

        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
        {
            if (turnContext.Activity.Name == "adaptiveCard/action")
            {
                var actionActivity = JsonConvert.DeserializeObject<AdaptiveCardInvokeValue>(turnContext.Activity.Value.ToString());
                string verb = actionActivity.Action.Verb;

                foreach (var handler in CardActionHandlers)
                {
                    if (handler.TriggerVerb.ToLowerInvariant() == verb?.ToString().ToLowerInvariant())
                    {
                        var invokeResponse = await handler.HandleActionInvokedAsync(turnContext, actionActivity.Action.Data, cancellationToken).ConfigureAwait(false);

                        if (invokeResponse != null)
                        {
                            var body = JObject.FromObject(invokeResponse.Body);
                            var responseValue = JsonConvert.DeserializeObject<AdaptiveCardInvokeResponse>(body.ToString());
                            string contentType = responseValue.Type;

                            if (contentType == InvokeResponseContentType.AdaptiveCard)
                            {
                                var card = responseValue.Value as string;
                                if (card == null)
                                {
                                    string errorMessage = "Adaptive card content cannot be null.";
                                    await SendInvokeResponseAsync(
                                        turnContext,
                                        InvokeResponseFactory.ErrorResponse(InvokeResponseErrorCode.InternalServerError,
                                        errorMessage), cancellationToken).ConfigureAwait(false);
                                    throw new ExceptionWithCode(errorMessage, ExceptionCode.InvalidParameter);
                                }

                                var isRefresh = JObject.Parse(card).ContainsKey("refresh");
                                var adaptiveCardResponse = handler.AdaptiveCardResponse;

                                if (isRefresh && handler.AdaptiveCardResponse == AdaptiveCardResponse.ReplaceForInteractor)
                                {
                                    // Ensure the base card for refresh action is able be viewed for all members
                                    // Otherwise, the refresh won't be triggered when user see the base card in Teams.
                                    adaptiveCardResponse = AdaptiveCardResponse.ReplaceForAll;
                                }

                                var messageActivity = MessageFactory.Attachment
                                (
                                    new Agents.Core.Models.Attachment
                                    {
                                        ContentType = "application/vnd.microsoft.card.adaptive",
                                        Content = card,
                                    }
                                );

                                switch (adaptiveCardResponse)
                                {
                                    case AdaptiveCardResponse.NewForAll:
                                        await SendInvokeResponseAsync(turnContext, InvokeResponseFactory.TextMessage(defaultMessage), cancellationToken).ConfigureAwait(false);
                                        await turnContext.SendActivityAsync(messageActivity, cancellationToken).ConfigureAwait(false);
                                        break;
                                    case AdaptiveCardResponse.ReplaceForAll:
                                        await SendInvokeResponseAsync(turnContext, invokeResponse, cancellationToken).ConfigureAwait(false);
                                        messageActivity.Id = turnContext.Activity.ReplyToId;
                                        await turnContext.UpdateActivityAsync(messageActivity, cancellationToken).ConfigureAwait(false);
                                        break;
                                    case AdaptiveCardResponse.ReplaceForInteractor:
                                        await SendInvokeResponseAsync(turnContext, invokeResponse, cancellationToken).ConfigureAwait(false);
                                        break;
                                }
                            }
                            else
                            {
                                // text message or error response
                                await SendInvokeResponseAsync(turnContext, invokeResponse, cancellationToken).ConfigureAwait(false);
                            }
                        }
 
                        break;
                    }
                }
            }

            await next(cancellationToken).ConfigureAwait(false);
        }

        private async static Task SendInvokeResponseAsync(ITurnContext context, InvokeResponse response, CancellationToken cancellationToken)
        {
            var invokeActivity = new Activity
            {
                Type = Microsoft.Bot.Schema.ActivityTypesEx.InvokeResponse,
                Value = response
            };

            await context.SendActivityAsync(invokeActivity, cancellationToken).ConfigureAwait(false);
        }
    }
}
