// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Storage, StoreItems, TurnContext, InvokeResponse } from "@microsoft/agents-hosting";
import { Activity } from "@microsoft/agents-activity";
import {
  AdaptiveCardResponse,
  CommandMessage,
  InvokeResponseErrorCode,
  MessageResponse,
  NotificationTarget,
  NotificationTargetType,
  TeamsFxAdaptiveCardActionHandler,
  TeamsFxBotCommandHandler,
  TeamsFxBotSsoCommandHandler,
  TriggerPatterns,
} from "../../../../src/conversation/interface";
import { InvokeResponseFactory } from "../../../../src/conversation/invokeResponseFactory";
import { TeamsBotSsoPromptTokenResponse } from "../../../../src";

export class TestTarget implements NotificationTarget {
  public content: any;
  public type?: NotificationTargetType | undefined;
  public error?: Error;
  public async sendMessage(
    text: string,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    if (this.error && onError) {
      await onError({} as TurnContext, this.error);
    }
    return new Promise((resolve) => {
      this.content = text;
      resolve({});
    });
  }
  public async sendAdaptiveCard(
    card: unknown,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    if (this.error && onError) {
      await onError({} as TurnContext, this.error);
    }
    return new Promise((resolve) => {
      this.content = card;
      resolve({});
    });
  }
}

export class TestSsoCommandHandler implements TeamsFxBotSsoCommandHandler {
  public triggerPatterns: TriggerPatterns;
  public responseMessage?: string | undefined;
  constructor(patterns: TriggerPatterns, responseMessage?: string) {
    this.triggerPatterns = patterns;
    if (responseMessage) {
      this.responseMessage = responseMessage;
    } else {
      this.responseMessage = "Sample command response";
    }
  }
  async handleCommandReceived(
    context: TurnContext,
    message: CommandMessage,
    ssoToken: TeamsBotSsoPromptTokenResponse
  ): Promise<string | void | Partial<Activity>> {
    return this.responseMessage;
  }
}

export class TestCommandHandler implements TeamsFxBotCommandHandler {
  public readonly triggerPatterns: TriggerPatterns;

  public isInvoked = false;
  public lastReceivedMessage: CommandMessage | undefined;

  constructor(patterns: TriggerPatterns) {
    this.triggerPatterns = patterns;
  }

  async handleCommandReceived(
    context: TurnContext,
    message: CommandMessage
  ): Promise<string | Partial<Activity> | void> {
    this.isInvoked = true;
    this.lastReceivedMessage = message;
    return "Sample command response";
  }
}

export class MockCardActionHandler implements TeamsFxAdaptiveCardActionHandler {
  isInvoked = false;
  triggerVerb: string;
  adaptiveCardResponse: AdaptiveCardResponse = AdaptiveCardResponse.ReplaceForInteractor;
  invokeResponse: InvokeResponse;
  actionData: any;

  constructor(verb: string, response?: any) {
    this.triggerVerb = verb;
    if (!response) {
      this.invokeResponse = InvokeResponseFactory.textMessage("Your response was sent to the app");
    } else if (typeof response === "string") {
      this.invokeResponse = InvokeResponseFactory.textMessage(response);
    } else {
      this.invokeResponse = InvokeResponseFactory.adaptiveCard(response);
    }
  }

  async handleActionInvoked(context: TurnContext, actionData: any): Promise<InvokeResponse> {
    this.isInvoked = true;
    this.actionData = actionData;
    return this.invokeResponse;
  }
}

export class MockCardActionHandlerWithErrorResponse implements TeamsFxAdaptiveCardActionHandler {
  isInvoked = false;
  triggerVerb: string;
  invokeResponse: InvokeResponse;
  actionData: any;

  constructor(verb: string, errorCode: InvokeResponseErrorCode, errorMessage: string) {
    this.triggerVerb = verb;
    this.invokeResponse = InvokeResponseFactory.errorResponse(errorCode, errorMessage);
  }

  async handleActionInvoked(context: TurnContext, actionData: any): Promise<InvokeResponse> {
    this.isInvoked = true;
    this.actionData = actionData;
    return this.invokeResponse;
  }
}

export class MockContext {
  private activity: any;
  constructor(text: string, type = "message") {
    this.activity = {
      text: text,
      type: type,
      recipient: {
        id: "1",
        name: "test-bot",
      },
    };
  }

  public sendActivity(activity: any): Promise<void> {
    return new Promise((resolve) => {
      resolve();
    });
  }
}

export class CustomStorage implements Storage {
  read(keys: string[]): Promise<StoreItems> {
    return Promise.resolve({});
  }
  write(changes: StoreItems): Promise<void> {
    return Promise.resolve();
  }
  delete(keys: string[]): Promise<void> {
    return Promise.resolve();
  }
}

export class MockActionInvokeContext {
  private activity: any;
  content: any;

  constructor(verb: string, data?: any) {
    this.activity = {
      type: "invoke",
      name: "adaptiveCard/action",
      value: {
        action: {
          type: "Action.Execute",
          verb: verb,
          data: data,
        },
      },
      trigger: "manual",
    };
  }

  public sendActivity(activity: any): Promise<void> {
    this.content = activity.value.body.value;
    return new Promise((resolve) => {
      resolve();
    });
  }
}
