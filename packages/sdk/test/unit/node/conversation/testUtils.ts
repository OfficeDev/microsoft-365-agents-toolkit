// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Storage,
  StoreItems,
  TurnContext,
  InvokeResponse,
  ResourceResponse,
  AttachmentData,
  AttachmentInfo,
  BaseAdapter,
} from "@microsoft/agents-hosting";
import {
  Activity,
  ActivityTypes,
  ConversationReference,
  ConversationAccount,
  ChannelAccount,
  Channels,
} from "@microsoft/agents-activity";
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
import assert from "assert";
import { v4 as uuidv4 } from "uuid";

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
      removeRecipientMention: () => {
        return text;
      },
      getConversationReference: () => {
        return {
          conversation: {
            id: "1",
          },
        };
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

  public updateActivity(activity: any): Promise<void> {
    return new Promise((resolve) => {
      resolve();
    });
  }
}

export class TestAdapter extends BaseAdapter {
  private logic: (context: TurnContext) => Promise<void>;
  readonly activeQueue: Partial<Activity>[] = [];

  constructor(logic: (context: TurnContext) => Promise<void>) {
    super();
    this.logic = logic;
  }

  public send(activity: string | Partial<Activity>): TestFlow {
    const messageActivity =
      typeof activity === "string" ? { type: ActivityTypes.Message, text: activity } : activity;
    return new TestFlow(this, this.processActivity(messageActivity));
  }

  public async processActivity(activity: Partial<Activity>): Promise<void> {
    const conversation = TestAdapter.createConversation(activity.text);
    const context = new TurnContext(this, {
      ...activity,
      conversation: conversation.conversation,
      user: conversation.user,
      from: conversation.user,
      getConversationReference: () => {
        return {};
      },
      removeRecipientMention: () => {
        return activity.text;
      },
    } as any);
    await this.runMiddleware(context, this.logic);
  }

  public async deleteActivity(context: TurnContext, reference: Partial<Activity>): Promise<void> {
    return Promise.resolve();
  }

  public async updateActivity(context: TurnContext, activity: Partial<Activity>): Promise<void> {
    return Promise.resolve();
  }

  public async continueConversation(
    reference: Partial<Activity>,
    logic: (context: TurnContext) => Promise<void>
  ): Promise<void> {
    return Promise.resolve();
  }

  public async sendActivities(
    context: TurnContext,
    activities: Partial<Activity>[]
  ): Promise<ResourceResponse[]> {
    if (!context) {
      throw new Error("TurnContext cannot be null.");
    }

    if (!activities) {
      throw new Error("Activities cannot be null.");
    }

    if (activities.length == 0) {
      throw new Error("Expecting one or more activities, but the array was empty.");
    }

    const responses: ResourceResponse[] = [];

    for (let i = 0; i < activities.length; i++) {
      const activity = activities[i];

      if (!activity.id) {
        activity.id = uuidv4();
      }

      if (!activity.timestamp) {
        activity.timestamp = new Date();
      }
      this.activeQueue.push(activity);

      responses.push({ id: activity.id } as ResourceResponse);
    }

    return responses;
  }

  public async uploadAttachment(
    conversationId: string,
    attachmentData: AttachmentData
  ): Promise<ResourceResponse> {
    return { id: "1" } as ResourceResponse;
  }

  public async getAttachmentInfo(attachmentId: string): Promise<AttachmentInfo> {
    return {} as AttachmentInfo;
  }

  public async getAttachment(attachmentId: string, viewId: string): Promise<NodeJS.ReadableStream> {
    return {} as NodeJS.ReadableStream;
  }

  static createConversation(name = "Hello", user = "User1", bot = "Bot"): ConversationReference {
    const conversationReference: ConversationReference = {
      channelId: Channels.Test,
      serviceUrl: "https://test.com",
      conversation: { isGroup: false, id: name, name: name } as ConversationAccount,
      user: { id: user.toLowerCase(), name: user } as ChannelAccount,
      agent: { id: bot.toLowerCase(), name: bot } as ChannelAccount,
      locale: "en-us",
    };
    return conversationReference;
  }
}

export type TestActivityInspector = (activity: Partial<Activity>, description?: string) => void;

export class TestFlow {
  private readonly adapter: TestAdapter;
  private readonly previous: Promise<void>;

  constructor(adapter: TestAdapter, previous: Promise<void>) {
    this.adapter = adapter;
    this.previous = previous;
  }

  public send(message: string | Partial<Activity>): TestFlow {
    return new TestFlow(
      this.adapter,
      this.previous.then(() =>
        this.adapter.processActivity(
          typeof message === "string" ? { type: ActivityTypes.Message, text: message } : message
        )
      )
    );
  }

  public assertReply(
    expected: string | Partial<Activity> | ((activity: Partial<Activity>) => void),
    description?: string
  ): TestFlow {
    function validateActivity(activity: Partial<Activity>, expected: Partial<Activity>): void {
      // tslint:disable-next-line:forin
      Object.keys(expected).forEach((prop: any) => {
        assert.equal((<any>activity)[prop], (<any>expected)[prop]);
      });
    }
    function defaultInspector(reply: Partial<Activity>, description2?: string): void {
      if (typeof expected === "object") {
        validateActivity(reply, expected);
      } else {
        assert.equal(
          reply.type,
          ActivityTypes.Message,
          `${description2} type === '${reply.type}'. `
        );
        assert.equal(reply.text, expected, `${description2} text === "${reply.text}"`);
      }
    }
    if (!description) {
      description = "";
    }
    const inspector: TestActivityInspector =
      typeof expected === "function" ? expected : defaultInspector;
    return new TestFlow(
      this.adapter,
      this.previous.then(() => {
        let timeout: number | undefined;
        // tslint:disable-next-line:promise-must-complete
        return new Promise<void>((resolve: any, reject: any): void => {
          if (!timeout) {
            timeout = 3000;
          }
          const start: number = new Date().getTime();
          const adapter: TestAdapter = this.adapter;

          function waitForActivity(): void {
            const current: number = new Date().getTime();
            if (current - start > <number>timeout) {
              // Operation timed out
              let expecting: string;
              switch (typeof expected) {
                case "string":
                default:
                  expecting = `"${expected.toString()}"`;
                  break;
                case "object":
                  expecting = `"${(expected as Activity).text}`;
                  break;
                case "function":
                  expecting = expected.toString();
                  break;
              }
              reject(
                new Error(
                  `TestAdapter.assertReply(${expecting}): ${description} Timed out after ${
                    current - start
                  }ms.`
                )
              );
            } else if (adapter.activeQueue.length > 0) {
              // Activity received
              const reply: Partial<Activity> = adapter.activeQueue.shift() as Activity;
              try {
                inspector(reply, description as string);
              } catch (err) {
                reject(err);
              }
              resolve();
            } else {
              setTimeout(waitForActivity, 5);
            }
          }
          waitForActivity();
        });
      })
    );
  }

  public catch(callback: (err: Error) => void): TestFlow {
    return new TestFlow(this.adapter, this.previous.catch(callback));
  }
}
