// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { InvokeResponse, StatusCodes } from "@microsoft/agents-hosting";
import { InvokeResponseErrorCode } from "./interface";

/**
 * Available response type for an adaptive card invoke response.
 * @internal
 */
export enum InvokeResponseType {
  AdaptiveCard = "application/vnd.microsoft.card.adaptive",
  Message = "application/vnd.microsoft.activity.message",
  Error = "application/vnd.microsoft.error",
}

/**
 * Provides methods for formatting various invoke responses a bot can send to respond to an invoke request.
 *
 * @remarks
 * All of these functions return an `InvokeResponse` object, which can be
 * passed as input to generate a new `invokeResponse` activity.
 *
 * This example sends an invoke response that contains an adaptive card.
 *
 * ```typescript
 *
 * const myCard = {
 *    type: "AdaptiveCard",
 *    body: [
 *     {
 *       "type": "TextBlock",
 *       "text": "This is a sample card"
 *     }],
 *     $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
 *     version: "1.4"
 *  };
 *
 * const invokeResponse = InvokeResponseFactory.adaptiveCard(myCard);
 * await context.sendActivity({
 *    type: ActivityTypes.InvokeResponse,
 *    value: invokeResponse,
 *  });
 * ```
 */
export class InvokeResponseFactory {
  /**
   * Create an invoke response from a text message.
   * The type of the invoke response is `application/vnd.microsoft.activity.message`
   * indicates the request was successfully processed.
   *
   * @param message - A text message included in a invoke response.
   *
   * @returns An `InvokeResponse` object.
   */
  public static textMessage(message: string): InvokeResponse {
    if (!message) {
      throw new Error("The text message cannot be null or empty");
    }

    return {
      status: StatusCodes.OK,
      body: {
        statusCode: StatusCodes.OK,
        type: InvokeResponseType.Message,
        value: message,
      },
    };
  }

  /**
   * Create an invoke response from an adaptive card.
   *
   * The type of the invoke response is `application/vnd.microsoft.card.adaptive` indicates
   * the request was successfully processed, and the response includes an adaptive card
   * that the client should display in place of the current one.
   *
   * @param card - The adaptive card JSON payload.
   *
   * @returns An `InvokeResponse` object.
   */
  public static adaptiveCard(card: unknown): InvokeResponse {
    if (!card) {
      throw new Error("The adaptive card content cannot be null or undefined");
    }

    return {
      status: StatusCodes.OK,
      body: {
        statusCode: StatusCodes.OK,
        type: InvokeResponseType.AdaptiveCard,
        value: card,
      },
    };
  }

  /**
   * Create an invoke response with error code and message.
   *
   * The type of the invoke response is `application/vnd.microsoft.error` indicates
   * the request was failed to processed.
   *
   * @param errorCode - The status code indicates error, available values:
   *  - 400 (BadRequest): indicate the incoming request was invalid.
   *  - 500 (InternalServerError): indicate an unexpected error occurred.
   * @param errorMessage - The error message.
   *
   * @returns An `InvokeResponse` object.
   */
  public static errorResponse(
    errorCode: InvokeResponseErrorCode,
    errorMessage: string
  ): InvokeResponse {
    return {
      status: StatusCodes.OK,
      body: {
        statusCode: errorCode,
        type: InvokeResponseType.Error,
        value: {
          code: errorCode.toString(),
          message: errorMessage,
        },
      },
    };
  }

  /**
   * Create an invoke response with status code and response value.
   * @param statusCode - The status code.
   * @param body - The value of the response body.
   *
   * @returns An `InvokeResponse` object.
   */
  public static createInvokeResponse(statusCode: StatusCodes, body?: unknown): InvokeResponse {
    return {
      status: statusCode,
      body: body,
    };
  }
}
