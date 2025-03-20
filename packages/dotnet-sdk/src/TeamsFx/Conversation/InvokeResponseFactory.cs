// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;
    using System.Net;
    using AdaptiveCardInvokeResponse = Agents.Core.Models.AdaptiveCardInvokeResponse;

    /// <summary>
    /// Contains utility methods for various invoke response types for an adaptiveCard/action invoke response.
    /// </summary>
    public static class InvokeResponseFactory
    {
        /// <summary>
        /// Creates invoke response with an adaptive card.
        /// </summary>
        /// <param name="card">The adaptive card included in the response body.</param>
        /// <returns>An instance of <see cref="InvokeResponse"/></returns>
        public static InvokeResponse AdaptiveCard(object card)
        {
            var response = new AdaptiveCardInvokeResponse()
            {
                StatusCode = 200,
                Type = InvokeResponseContentType.AdaptiveCard,
                Value = card
            };

            return new InvokeResponse
            {
                Status = (int)HttpStatusCode.OK,
                Body = response
            };
        }

        /// <summary>
        /// Creates invoke response with a text message.
        /// </summary>
        /// <param name="message">The text message included in the response body.</param>
        /// <returns>An instance of <see cref="InvokeResponse"/>.</returns>
        public static InvokeResponse TextMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                throw new ArgumentNullException(nameof(message));
            }

            var response = new AdaptiveCardInvokeResponse()
            {
                StatusCode = 200,
                Type = InvokeResponseContentType.Message,
                Value = message
            };

            return new InvokeResponse
            {
                Status = (int)HttpStatusCode.OK,
                Body = response
            };
        }

        /// <summary>
        /// Creates invoke response with error code and error message.
        /// </summary>
        /// <param name="errorCode">The status code indicates the bot processing error, available values:
        /// <list type="bullet">
        ///     <item>
        ///         <description> 400 (BadRequest): indicate the incoming request was invalid.</description>
        ///     </item>
        ///     <item>
        ///         <description> 500 (InternalServerError): indicate an unexpected error occurred.</description>
        ///     </item>
        /// </list>
        /// </param>
        /// <param name="errorMessage">The error message.</param>
        /// <returns>An instance of <see cref="InvokeResponse"/></returns>
        public static InvokeResponse ErrorResponse(InvokeResponseErrorCode errorCode, string errorMessage)
        {
            if (string.IsNullOrWhiteSpace(errorMessage))
            {
                throw new ArgumentNullException(nameof(errorMessage));
            }

            var response = new AdaptiveCardInvokeResponse()
            {
                StatusCode = (int)errorCode,
                Type = InvokeResponseContentType.Error,
                Value = new Error()
                {
                    Code = errorCode.ToString(),
                    Message = errorMessage
                }
            };

            return new InvokeResponse
            {
                Status = (int)HttpStatusCode.OK,
                Body = response
            };
        }

        /// <summary>
        /// Creates invoke response with status code and response body.
        /// </summary>
        /// <param name="statusCode">The status code.</param>
        /// <param name="response">The response body.</param>
        /// <returns>An instance of <see cref="InvokeResponse"/>.</returns>
        public static InvokeResponse CreateInvokeResponse(HttpStatusCode statusCode, object response)
        {
            return new InvokeResponse
            {
                Status = (int)statusCode,
                Body = response
            };
        }
    }
}
