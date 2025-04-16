// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ActivityTypes, Activity } from "@microsoft/agents-activity";
import {
  CardFactory,
  TurnContext,
  InvokeResponse,
  MessageFactory,
  Middleware,
} from "@microsoft/agents-hosting";
import {
  AdaptiveCardResponse,
  InvokeResponseErrorCode,
  TeamsFxAdaptiveCardActionHandler,
} from "../interface";
import { InvokeResponseFactory, InvokeResponseType } from "../invokeResponseFactory";

/**
 * @internal
 */
export class CardActionMiddleware implements Middleware {
  public readonly actionHandlers: TeamsFxAdaptiveCardActionHandler[] = [];
  private readonly defaultMessage: string = "Your response was sent to the app";

  constructor(handlers?: TeamsFxAdaptiveCardActionHandler[]) {
    if (handlers && handlers.length > 0) {
      this.actionHandlers.push(...handlers);
    }
  }

  async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
    if (context.activity.name === "adaptiveCard/action") {
      const action = (context.activity.value as any).action;
      const actionVerb = action.verb;

      for (const handler of this.actionHandlers) {
        if (handler.triggerVerb?.toLowerCase() === actionVerb?.toLowerCase()) {
          let response: InvokeResponse;
          try {
            response = await handler.handleActionInvoked(context, action.data);
          } catch (error: any) {
            const errorResponse = InvokeResponseFactory.errorResponse(
              InvokeResponseErrorCode.InternalServerError,
              error.message
            );
            await this.sendInvokeResponse(context, errorResponse);
            throw error;
          }

          const responseType = response.body?.type;
          switch (responseType) {
            case InvokeResponseType.AdaptiveCard:
              const card = response.body?.value;
              if (!card) {
                const errorMessage = "Adaptive card content cannot be found in the response body";
                await this.sendInvokeResponse(
                  context,
                  InvokeResponseFactory.errorResponse(
                    InvokeResponseErrorCode.InternalServerError,
                    errorMessage
                  )
                );
                throw new Error(errorMessage);
              }

              if (card.refresh && handler.adaptiveCardResponse !== AdaptiveCardResponse.NewForAll) {
                // Card won't be refreshed with AdaptiveCardResponse.ReplaceForInteractor.
                // So set to AdaptiveCardResponse.ReplaceForAll here.
                handler.adaptiveCardResponse = AdaptiveCardResponse.ReplaceForAll;
              }

              const activity = MessageFactory.attachment(CardFactory.adaptiveCard(card));
              if (handler.adaptiveCardResponse === AdaptiveCardResponse.NewForAll) {
                await this.sendInvokeResponse(
                  context,
                  InvokeResponseFactory.textMessage(this.defaultMessage)
                );
                await context.sendActivity(activity);
              } else if (handler.adaptiveCardResponse === AdaptiveCardResponse.ReplaceForAll) {
                activity.id = context.activity.replyToId;
                await context.updateActivity(activity);
                await this.sendInvokeResponse(context, response);
              } else {
                await this.sendInvokeResponse(context, response);
              }
              break;
            case InvokeResponseType.Message:
            case InvokeResponseType.Error:
            default:
              await this.sendInvokeResponse(context, response);
              break;
          }

          break;
        }
      }
    }

    await next();
  }

  private async sendInvokeResponse(context: TurnContext, response: InvokeResponse) {
    await context.sendActivity({
      type: ActivityTypes.InvokeResponse,
      value: response,
    } as Activity);
  }
}
