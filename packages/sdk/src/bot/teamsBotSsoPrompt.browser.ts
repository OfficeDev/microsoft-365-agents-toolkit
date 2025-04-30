// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DialogContext, DialogTurnResult } from "@microsoft/agents-hosting-dialogs";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";
import { OnBehalfOfCredentialAuthConfig } from "../models/configuration";

/**
 * Settings used to configure an TeamsBotSsoPrompt instance.
 *
 * @remarks
 * Only works in in server side.
 */
export interface TeamsBotSsoPromptSettings {
  /**
   * The array of strings that declare the desired permissions and the resources requested.
   */
  scopes: string[];

  /**
   * (Optional) number of milliseconds the prompt will wait for the user to authenticate.
   * Defaults to a value `900,000` (15 minutes.)
   */
  timeout?: number;

  /**
   * (Optional) value indicating whether the TeamsBotSsoPrompt should end upon receiving an
   * invalid message.  Generally the TeamsBotSsoPrompt will end the auth flow when receives user
   * message not related to the auth flow. Setting the flag to false ignores the user's message instead.
   * Defaults to value `true`
   */
  endOnInvalidMessage?: boolean;
}

/**
 * Creates a new prompt that leverage Teams Single Sign On (SSO) support for bot to automatically sign in user and
 * help receive oauth token, asks the user to consent if needed.
 *
 * @remarks
 * The prompt will attempt to retrieve the users current token of the desired scopes and store it in
 * the token store.
 *
 * User will be automatically signed in leveraging Teams support of Bot Single Sign On(SSO):
 * https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-aad-sso-bots
 *
 * @example
 * When used with your bots `DialogSet` you can simply add a new instance of the prompt as a named
 * dialog using `DialogSet.add()`. You can then start the prompt from a waterfall step using either
 * `DialogContext.beginDialog()` or `DialogContext.prompt()`. The user will be prompted to sign in as
 * needed and their access token will be passed as an argument to the callers next waterfall step:
 *
 * ```JavaScript
 * const { ConversationState, MemoryStorage } = require('botbuilder');
 * const { DialogSet, WaterfallDialog } = require('botbuilder-dialogs');
 * const { TeamsBotSsoPrompt } = require('@microsoft/teamsfx');
 *
 * const convoState = new ConversationState(new MemoryStorage());
 * const dialogState = convoState.createProperty('dialogState');
 * const dialogs = new DialogSet(dialogState);
 *
 * dialogs.add(new TeamsBotSsoPrompt('TeamsBotSsoPrompt', {
 *    scopes: ["User.Read"],
 * }));
 *
 * dialogs.add(new WaterfallDialog('taskNeedingLogin', [
 *      async (step) => {
 *          return await step.beginDialog('TeamsBotSsoPrompt');
 *      },
 *      async (step) => {
 *          const token = step.result;
 *          if (token) {
 *
 *              // ... continue with task needing access token ...
 *
 *          } else {
 *              await step.context.sendActivity(`Sorry... We couldn't log you in. Try again later.`);
 *              return await step.endDialog();
 *          }
 *      }
 * ]));
 * ```
 */
export class TeamsBotSsoPrompt {
  /**
   * Constructor of TeamsBotSsoPrompt.
   *
   * @param dialogId Unique ID of the dialog within its parent `DialogSet` or `ComponentDialog`.
   * @param settings Settings used to configure the prompt.
   *
   * @throws {@link ErrorCode|InvalidParameter} when scopes is not a valid string or string array.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   */
  constructor(
    authConfig: OnBehalfOfCredentialAuthConfig,
    initiateLoginEndpoint: string,
    dialogId: string,
    settings: TeamsBotSsoPromptSettings
  ) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotSsoPrompt"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Called when a prompt dialog is pushed onto the dialog stack and is being activated.
   * @remarks
   * If the task is successful, the result indicates whether the prompt is still
   * active after the turn has been processed by the prompt.
   *
   * @param dc The DialogContext for the current turn of the conversation.
   *
   * @throws {@link ErrorCode|InvalidParameter} when timeout property in teams bot sso prompt settings is not number or is not positive.
   * @throws {@link ErrorCode|ChannelNotSupported} when bot channel is not MS Teams.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   *
   * @returns A `Promise` representing the asynchronous operation.
   */
  public beginDialog(dc: DialogContext): Promise<DialogTurnResult> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotSsoPrompt"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Called when a prompt dialog is the active dialog and the user replied with a new activity.
   *
   * @remarks
   * If the task is successful, the result indicates whether the dialog is still
   * active after the turn has been processed by the dialog.
   * The prompt generally continues to receive the user's replies until it accepts the
   * user's reply as valid input for the prompt.
   *
   * @param dc The DialogContext for the current turn of the conversation.
   *
   * @returns A `Promise` representing the asynchronous operation.
   *
   * @throws {@link ErrorCode|ChannelNotSupported} when bot channel is not MS Teams.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   */
  public continueDialog(dc: DialogContext): Promise<DialogTurnResult> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotSsoPrompt"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}
