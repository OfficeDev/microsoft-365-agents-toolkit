// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConversationState,
  UserState,
  TurnContext,
  InvokeResponse,
  Storage,
} from "@microsoft/agents-hosting";
import { Activity, ConversationReference } from "@microsoft/agents-activity";
import { SigninStateVerificationQuery } from "@microsoft/agents-hosting-teams";
import { TeamsBotSsoPromptTokenResponse } from "../bot/teamsBotSsoPromptTokenResponse";
import { OnBehalfOfCredentialAuthConfig } from "../models/configuration";

/**
 * The response of a message action, e.g., `sendMessage`, `sendAdaptiveCard`.
 */
export interface MessageResponse {
  /**
   * Id of the message.
   */
  id?: string;
}

/**
 * The target type where the notification will be sent to.
 *
 * @remarks
 * - "Channel" means to a team channel. (By default, notification to a team will be sent to its "General" channel.)
 * - "Group" means to a group chat.
 * - "Person" means to a personal chat.
 */
export enum NotificationTargetType {
  /**
   * The notification will be sent to a team channel.
   * (By default, notification to a team will be sent to its "General" channel.)
   */
  Channel = "Channel",
  /**
   * The notification will be sent to a group chat.
   */
  Group = "Group",
  /**
   * The notification will be sent to a personal chat.
   */
  Person = "Person",
}

/**
 * Represent a notification target.
 */
export interface NotificationTarget {
  /**
   * The type of target, could be "Channel" or "Group" or "Person".
   */
  readonly type?: NotificationTargetType;

  /**
   * Send a plain text message.
   *
   * @param text - the plain text message.
   * @param onError - an optional error handler that can catch exceptions during message sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns the response of sending message.
   */
  sendMessage(
    text: string,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse>;

  /**
   * Send an adaptive card message.
   *
   * @param card - the adaptive card raw JSON.
   * @param onError - an optional error handler that can catch exceptions during adaptive card sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns the response of sending adaptive card message.
   */
  sendAdaptiveCard(
    card: unknown,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse>;
}

/**
 * A store to persist notification target references.
 */
export interface ConversationReferenceStore {
  /**
   * Add a conversation reference to the store. If overwrite, update existing one, otherwise add when not exist.
   *
   * @param key the key of the conversation reference.
   * @param reference the conversation reference to add.
   * @param options the options to add the conversation reference.
   *
   * @returns true if added or updated, false if not changed.
   */
  add(
    key: string,
    reference: Partial<ConversationReference>,
    options: ConversationReferenceStoreAddOptions
  ): Promise<boolean>;

  /**
   * Remove a conversation reference from the store.
   *
   * @param key the key of the conversation reference.
   * @param reference the conversation reference to remove.
   *
   * @returns true if exist and removed, false if not changed.
   */
  remove(key: string, reference: Partial<ConversationReference>): Promise<boolean>;

  /**
   * List stored conversation reference by page.
   *
   * @param pageSize the page size.
   * @param continuationToken the continuation token to get next page.
   *
   * @returns a paged list of conversation references.
   */
  list(
    pageSize?: number,
    continuationToken?: string
  ): Promise<PagedData<Partial<ConversationReference>>>;
}

/**
 * Options to add a conversation reference to the store.
 */
export interface ConversationReferenceStoreAddOptions {
  /**
   * Whether to overwrite the existing conversation reference.
   */
  overwrite?: boolean;
}

/**
 * Represents a page of data.
 */
export interface PagedData<T> {
  /**
   * Page of data.
   */
  data: T[];

  /**
   * The Continuation Token to pass to get the next page of results.
   *
   * @remarks
   * Undefined or empty token means the page reaches the end.
   */
  continuationToken?: string;
}

/**
 * The trigger pattern used to trigger a {@link TeamsFxBotCommandHandler} instance.
 */
export type TriggerPatterns = string | RegExp | (string | RegExp)[];

/**
 * Interface for a command message that can handled in a command handler.
 */
export interface CommandMessage {
  /**
   * Text of the message sent by the user.
   */
  text: string;

  /**
   * The capture groups that matched to the {@link TriggerPatterns} in a {@link TeamsFxBotCommandHandler} instance.
   */
  matches?: RegExpMatchArray;
}

/**
 * Interface for a command handler that can process command to a TeamsFx bot and return a response.
 */
export interface TeamsFxBotCommandHandler {
  /**
   * The string or regular expression patterns that can trigger this handler.
   */
  triggerPatterns: TriggerPatterns;

  /**
   * Handles a bot command received activity.
   *
   * @param context The bot context.
   * @param message The command message the user types from Teams.
   * @returns A `Promise` representing an activity or text to send as the command response.
   * Or no return value if developers want to send the response activity by themselves in this method.
   */
  handleCommandReceived(
    context: TurnContext,
    message: CommandMessage
  ): Promise<string | Partial<Activity> | void>;
}

/**
 * Interface for a command handler that can process sso command to a TeamsFx bot and return a response.
 */
export interface TeamsFxBotSsoCommandHandler {
  /**
   * The string or regular expression patterns that can trigger this handler.
   */
  triggerPatterns: TriggerPatterns;

  /**
   * Handles a bot command received activity.
   *
   * @param context The bot context.
   * @param message The command message the user types from Teams.
   * @param tokenResponse The tokenResponse which contains sso token that can be used to exchange access token for the bot.
   * @returns A `Promise` representing an activity or text to send as the command response.
   * Or no return value if developers want to send the response activity by themselves in this method.
   */
  handleCommandReceived(
    context: TurnContext,
    message: CommandMessage,
    tokenResponse: TeamsBotSsoPromptTokenResponse
  ): Promise<string | Partial<Activity> | void>;
}

/**
 * Options to initialize {@link CommandBot}.
 */
export interface CommandOptions {
  /**
   * The commands to registered with the command bot. Each command should implement the interface {@link TeamsFxBotCommandHandler} so that it can be correctly handled by this command bot.
   */
  commands?: TeamsFxBotCommandHandler[];

  /**
   * The commands to registered with the sso command bot. Each sso command should implement the interface {@link TeamsFxBotSsoCommandHandler} so that it can be correctly handled by this command bot.
   */
  ssoCommands?: TeamsFxBotSsoCommandHandler[];
}

/**
 * Options to initialize {@link CardActionBot}.
 */
export interface CardActionOptions {
  /**
   * The action handlers to registered with the action bot. Each command should implement the interface {@link TeamsFxAdaptiveCardActionHandler} so that it can be correctly handled by this bot.
   */
  actions?: TeamsFxAdaptiveCardActionHandler[];
}

/**
 * Options used to control how the response card will be sent to users.
 */
export enum AdaptiveCardResponse {
  /**
   * The response card will be replaced the current one for the interactor who trigger the action.
   */
  ReplaceForInteractor,

  /**
   * The response card will be replaced the current one for all users in the chat.
   */
  ReplaceForAll,

  /**
   * The response card will be sent as a new message for all users in the chat.
   */
  NewForAll,
}

/**
 * Status code for an `application/vnd.microsoft.error` invoke response.
 */
export enum InvokeResponseErrorCode {
  /**
   * Invalid request.
   */
  BadRequest = 400,

  /**
   * Internal server error.
   */
  InternalServerError = 500,
}

/**
 * Interface for adaptive card action handler that can process card action invoke and return a response.
 */
export interface TeamsFxAdaptiveCardActionHandler {
  /**
   * The verb defined in adaptive card action that can trigger this handler.
   * The verb string here is case-insensitive.
   */
  triggerVerb: string;

  /**
   * Specify the behavior for how the card response will be sent in Teams conversation.
   * The default value is `AdaptiveCardResponse.ReplaceForInteractor`, which means the card
   * response will replace the current one only for the interactor.
   */
  adaptiveCardResponse?: AdaptiveCardResponse;

  /**
   * The handler function that will be invoked when the action is fired.
   * @param context The turn context.
   * @param actionData The contextual data that associated with the action.
   * 
   * @returns A `Promise` representing a invoke response for the adaptive card invoke action.
   * You can use the `InvokeResponseFactory` utility class to create an invoke response from
   *  - A text message: 
   *   ```typescript 
   *   return InvokeResponseFactory.textMessage("Action is processed successfully!");
   *   ```
   *  - An adaptive card:
   *    ```typescript
   *    const responseCard = AdaptiveCards.declare(helloWorldCard).render(actionData);
        return InvokeResponseFactory.adaptiveCard(responseCard);
   *    ```
   *  - An error response:
   *     ```typescript
   *     return InvokeResponseFactory.errorResponse(InvokeResponseErrorCode.BadRequest, "Invalid request");
   *     ```
   * 
   * @remarks For more details about the invoke response format, refer to https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/universal-action-model#response-format.
   */
  handleActionInvoked(context: TurnContext, actionData: any): Promise<InvokeResponse>;
}

/**
 * Interface for SSO configuration for Bot SSO
 */
export interface BotSsoConfig {
  /**
   * aad related configurations
   */
  aad: {
    /**
     * The list of scopes for which the token will have access
     */
    scopes: string[];
  } & OnBehalfOfCredentialAuthConfig & { initiateLoginEndpoint: string };

  dialog?: {
    // eslint-disable-next-line no-secrets/no-secrets
    /**
     * Custom sso execution activity handler class which should implement the interface {@link BotSsoExecutionActivityHandler}. If not provided, it will use {@link DefaultBotSsoExecutionActivityHandler} by default
     */
    CustomBotSsoExecutionActivityHandler?: new (
      ssoConfig: BotSsoConfig
    ) => BotSsoExecutionActivityHandler;

    /**
     * Conversation state for sso command bot, if not provided, it will use internal memory storage to create a new one.
     */
    conversationState?: ConversationState;

    /**
     * User state for sso command bot, if not provided, it will use internal memory storage to create a new one.
     */
    userState?: UserState;

    /**
     * Used by {@link BotSsoExecutionDialog} to remove duplicated messages, if not provided, it will use internal memory storage
     */
    dedupStorage?: Storage;

    /**
     * Settings used to configure an teams sso prompt dialog.
     */
    ssoPromptConfig?: {
      /**
       * Number of milliseconds the prompt will wait for the user to authenticate.
       * Defaults to a value `900,000` (15 minutes.)
       */
      timeout?: number;

      /**
       * Value indicating whether the TeamsBotSsoPrompt should end upon receiving an
       * invalid message.  Generally the TeamsBotSsoPrompt will end the auth flow when receives user
       * message not related to the auth flow. Setting the flag to false ignores the user's message instead.
       * Defaults to value `true`
       */
      endOnInvalidMessage?: boolean;
    };
  };
}

// eslint-disable-next-line no-secrets/no-secrets
/**
 * Interface for user to customize SSO execution activity handler
 *
 * @remarks
 * Bot SSO execution activity handler is to handle SSO login process and trigger SSO command using {@link BotSsoExecutionDialog}.
 * You can use this interface to implement your own SSO execution dialog, and pass it to ConversationBot options:
 *
 * ```typescript
 * export const commandBot = new ConversationBot({
 *   ...
 *   ssoConfig: {
 *     ...
 *     dialog: {
 *       CustomBotSsoExecutionActivityHandler: YourCustomBotSsoExecutionActivityHandler,
 *     }
 *   },
 *    ...
 * });
 * ```
 * For details information about how to implement a BotSsoExecutionActivityHandler, please refer DefaultBotSsoExecutionActivityHandler class source code: https://aka.ms/teamsfx-default-sso-execution-activity-handler
 */
export interface BotSsoExecutionActivityHandler {
  // eslint-disable-next-line no-secrets/no-secrets
  /**
   * Add {@link TeamsFxBotSsoCommandHandler} instance to {@link BotSsoExecutionDialog}
   * @param handler {@link BotSsoExecutionDialogHandler} callback function
   * @param triggerPatterns The trigger pattern
   *
   * @remarks
   * This function is used to add SSO command to {@link BotSsoExecutionDialog} instance.
   */
  addCommand(handler: BotSsoExecutionDialogHandler, triggerPatterns: TriggerPatterns): void;

  /**
   * Called to initiate the event emission process.
   * @param context The context object for the current turn.
   */
  run(context: TurnContext): Promise<void>;

  /**
   * Receives invoke activities with Activity name of 'signin/verifyState'.
   * @param context A context object for this turn.
   * @param query Signin state (part of signin action auth flow) verification invoke query.
   * @returns A promise that represents the work queued.
   *
   * @remarks
   * It should trigger {@link BotSsoExecutionDialog} instance to handle signin process
   */
  handleTeamsSigninVerifyState(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ): Promise<void>;

  /**
   * Receives invoke activities with Activity name of 'signin/tokenExchange'
   * @param context A context object for this turn.
   * @param query Signin state (part of signin action auth flow) verification invoke query
   * @returns A promise that represents the work queued.
   *
   * @remarks
   * It should trigger {@link BotSsoExecutionDialog} instance to handle signin process
   */
  handleTeamsSigninTokenExchange(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ): Promise<void>;
}

export type BotSsoExecutionDialogHandler = (
  context: TurnContext,
  tokenResponse: TeamsBotSsoPromptTokenResponse,
  message: CommandMessage
) => Promise<void>;
