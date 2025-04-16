// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken } from "@azure/identity";
import {
  CardFactory,
  MessageFactory,
  TurnContext,
  OAuthCard,
  StatusCodes,
  TokenExchangeInvokeRequest,
  TokenExchangeResource,
  TokenPostResource,
  SigningResource,
} from "@microsoft/agents-hosting";
import {
  TeamsInfo,
  TeamsChannelAccount,
  verifyStateOperationName,
  tokenExchangeOperationName,
} from "@microsoft/agents-hosting-teams";
import {
  Dialog,
  DialogContext,
  DialogTurnResult,
  PromptOptions,
  PromptRecognizerResult,
} from "@microsoft/agents-hosting-dialogs";
import { Activity, ActivityTypes, Channels, ActionTypes } from "@microsoft/agents-activity";
import { TeamsBotSsoPromptTokenResponse } from "./teamsBotSsoPromptTokenResponse";
import { v4 as uuidv4 } from "uuid";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { internalLogger } from "../util/logger";
import { validateScopesType, formatString, parseJwt, validateConfig } from "../util/utils";
import { OnBehalfOfCredentialAuthConfig } from "../models/configuration";
import { OnBehalfOfUserCredential } from "../credential/onBehalfOfUserCredential";

const invokeResponseType = "invokeResponse";
/**
 * Response body returned for a token exchange invoke activity.
 */
class TokenExchangeInvokeResponse {
  /**
   * Response id
   */
  id: string;

  /**
   * Detailed error message
   */
  failureDetail: string;

  constructor(id: string, failureDetail: string) {
    this.id = id;
    this.failureDetail = failureDetail;
  }
}

/**
 * Settings used to configure an TeamsBotSsoPrompt instance.
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
export class TeamsBotSsoPrompt extends Dialog {
  private authConfig: OnBehalfOfCredentialAuthConfig;
  private initiateLoginEndpoint: string;
  private settings: TeamsBotSsoPromptSettings;
  /**
   * Constructor of TeamsBotSsoPrompt.
   *
   * @param {OnBehalfOfCredentialAuthConfig} authConfig - Used to provide configuration and auth
   * @param {string} initiateLoginEndpoint - Login URL for Teams to redirect to
   * @param {string} dialogId Unique ID of the dialog within its parent `DialogSet` or `ComponentDialog`.
   * @param {TeamsBotSsoPromptSettings} settings Settings used to configure the prompt.
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
    super(dialogId);

    this.initiateLoginEndpoint = initiateLoginEndpoint;
    this.authConfig = authConfig;
    this.settings = settings;

    validateScopesType(this.settings.scopes);
    validateConfig(this.authConfig);

    internalLogger.info("Create a new Teams Bot SSO Prompt");
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
  public async beginDialog(dc: DialogContext): Promise<DialogTurnResult> {
    internalLogger.info("Begin Teams Bot SSO Prompt");
    this.ensureMsTeamsChannel(dc);

    // Initialize prompt state
    const default_timeout = 900000;
    let timeout: number = default_timeout;
    if (this.settings.timeout) {
      if (typeof this.settings.timeout != "number") {
        const errorMsg = "type of timeout property in teamsBotSsoPromptSettings should be number.";
        internalLogger.error(errorMsg);
        throw new ErrorWithCode(errorMsg, ErrorCode.InvalidParameter);
      }
      if (this.settings.timeout <= 0) {
        const errorMsg =
          "value of timeout property in teamsBotSsoPromptSettings should be positive.";
        internalLogger.error(errorMsg);
        throw new ErrorWithCode(errorMsg, ErrorCode.InvalidParameter);
      }
      timeout = this.settings.timeout;
    }

    if (this.settings.endOnInvalidMessage === undefined) {
      this.settings.endOnInvalidMessage = true;
    }
    const state: teamsBotSsoPromptState = dc.activeDialog?.state as teamsBotSsoPromptState;
    state.state = {};
    state.options = {};
    state.expires = new Date().getTime() + timeout;

    // Send OAuth card to get SSO token
    await this.sendOAuthCardAsync(dc.context);
    return Dialog.EndOfTurn;
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
  public async continueDialog(dc: DialogContext): Promise<DialogTurnResult> {
    internalLogger.info("Continue Teams Bot SSO Prompt");
    this.ensureMsTeamsChannel(dc);

    // Check for timeout
    const state: teamsBotSsoPromptState = dc.activeDialog?.state as teamsBotSsoPromptState;
    const isMessage: boolean = dc.context.activity.type === ActivityTypes.Message;
    const isTimeoutActivityType: boolean =
      isMessage ||
      this.isTeamsVerificationInvoke(dc.context) ||
      this.isTokenExchangeRequestInvoke(dc.context);

    // If the incoming Activity is a message, or an Activity Type normally handled by TeamsBotSsoPrompt,
    // check to see if this TeamsBotSsoPrompt Expiration has elapsed, and end the dialog if so.
    const hasTimedOut: boolean = isTimeoutActivityType && new Date().getTime() > state.expires;
    if (hasTimedOut) {
      internalLogger.warn("End Teams Bot SSO Prompt due to timeout");
      return await dc.endDialog(undefined);
    } else {
      if (
        this.isTeamsVerificationInvoke(dc.context) ||
        this.isTokenExchangeRequestInvoke(dc.context)
      ) {
        // Recognize token
        const recognized: PromptRecognizerResult<TeamsBotSsoPromptTokenResponse> =
          await this.recognizeToken(dc);

        if (recognized.succeeded) {
          return await dc.endDialog(recognized.value);
        }
      } else if (isMessage && this.settings.endOnInvalidMessage) {
        internalLogger.warn("End Teams Bot SSO Prompt due to invalid message");
        return await dc.endDialog(undefined);
      }

      return Dialog.EndOfTurn;
    }
  }

  /**
   * Ensure bot is running in MS Teams since TeamsBotSsoPrompt is only supported in MS Teams channel.
   * @param dc dialog context
   * @throws {@link ErrorCode|ChannelNotSupported} if bot channel is not MS Teams
   * @internal
   */
  private ensureMsTeamsChannel(dc: DialogContext) {
    if (dc.context.activity.channelId != Channels.Msteams) {
      const errorMsg = formatString(
        ErrorMessage.OnlyMSTeamsChannelSupported,
        "Teams Bot SSO Prompt"
      );
      internalLogger.error(errorMsg);
      throw new ErrorWithCode(errorMsg, ErrorCode.ChannelNotSupported);
    }
  }

  /**
   * Send OAuthCard that tells Teams to obtain an authentication token for the bot application.
   * For details see https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-aad-sso-bots.
   *
   * @internal
   */
  private async sendOAuthCardAsync(context: TurnContext): Promise<void> {
    internalLogger.verbose("Send OAuth card to get SSO token");

    const account: TeamsChannelAccount = await TeamsInfo.getMember(
      context,
      context.activity.from!.id!
    );
    internalLogger.verbose(
      "Get Teams member account user principal name: " +
        (account.userPrincipalName ? account.userPrincipalName : "")
    );

    const loginHint: string = account.userPrincipalName ? account.userPrincipalName : "";
    const signInResource = this.getSignInResource(loginHint);
    const card = CardFactory.oauthCard("", "Teams SSO Sign In", "Sign In", signInResource);
    (card.content as OAuthCard).buttons[0].type = ActionTypes.Signin;
    const msg: Partial<Activity> = MessageFactory.attachment(card);

    // Send prompt
    await context.sendActivity(msg as Activity);
  }

  /**
   * Get sign in resource.
   *
   * @throws {@link ErrorCode|InvalidConfiguration} if client id, tenant id or initiate login endpoint is not found in config.
   *
   * @internal
   */
  private getSignInResource(loginHint: string): SigningResource {
    internalLogger.verbose("Get sign in authentication configuration");

    const signInLink = `${this.initiateLoginEndpoint}?scope=${encodeURI(
      this.settings.scopes.join(" ")
    )}&clientId=${this.authConfig.clientId}&tenantId=${
      this.authConfig.tenantId
    }&loginHint=${loginHint}`;

    internalLogger.verbose("Sign in link: " + signInLink);

    const tokenExchangeResource: TokenExchangeResource = {
      id: uuidv4(),
    };

    const tokenPostResource: TokenPostResource = {};

    return {
      signingLink: signInLink,
      tokenExchangeResource: tokenExchangeResource,
      tokenPostResource: tokenPostResource,
    };
  }

  /**
   * @internal
   */
  private async recognizeToken(
    dc: DialogContext
  ): Promise<PromptRecognizerResult<TeamsBotSsoPromptTokenResponse>> {
    const context = dc.context;
    let tokenResponse: TeamsBotSsoPromptTokenResponse | undefined;

    if (this.isTokenExchangeRequestInvoke(context)) {
      internalLogger.verbose("Receive token exchange request");
      // Received activity is not a token exchange request
      if (!(context.activity.value && this.isTokenExchangeRequest(context.activity.value))) {
        const warningMsg =
          "The bot received an InvokeActivity that is missing a TokenExchangeInvokeRequest value. This is required to be sent with the InvokeActivity.";

        internalLogger.warn(warningMsg);
        await context.sendActivity(
          this.getTokenExchangeInvokeResponse(StatusCodes.BAD_REQUEST, warningMsg)
        );
      } else {
        const ssoToken = context.activity.value.token;
        const credential = new OnBehalfOfUserCredential(ssoToken, this.authConfig);
        let exchangedToken: AccessToken | null;
        try {
          exchangedToken = await credential.getToken(this.settings.scopes);

          if (exchangedToken) {
            await context.sendActivity(
              this.getTokenExchangeInvokeResponse(StatusCodes.OK, "", context.activity.value.id)
            );

            const ssoTokenExpiration = parseJwt(ssoToken).exp;
            tokenResponse = {
              ssoToken: ssoToken,
              ssoTokenExpiration: new Date(ssoTokenExpiration * 1000).toISOString(),
              connectionName: "",
              token: exchangedToken.token,
              expiration: exchangedToken.expiresOnTimestamp.toString(),
            };
          }
        } catch (error) {
          const warningMsg = "The bot is unable to exchange token. Ask for user consent.";
          internalLogger.info(warningMsg);
          await context.sendActivity(
            this.getTokenExchangeInvokeResponse(
              StatusCodes.PRECONDITION_FAILED,
              warningMsg,
              context.activity.value.id
            )
          );
        }
      }
    } else if (this.isTeamsVerificationInvoke(context)) {
      internalLogger.verbose("Receive Teams state verification request");
      await this.sendOAuthCardAsync(dc.context);
      await context.sendActivity({
        type: invokeResponseType,
        value: { status: StatusCodes.OK },
      } as Activity);
    }

    return tokenResponse !== undefined
      ? { succeeded: true, value: tokenResponse }
      : { succeeded: false };
  }

  /**
   * @internal
   */
  private getTokenExchangeInvokeResponse(
    status: number,
    failureDetail: string,
    id?: string
  ): Activity {
    const invokeResponse: Partial<Activity> = {
      type: invokeResponseType,
      value: { status, body: new TokenExchangeInvokeResponse(id as string, failureDetail) },
    };
    return invokeResponse as Activity;
  }

  /**
   * @internal
   */
  private isTeamsVerificationInvoke(context: TurnContext): boolean {
    const activity: Activity = context.activity;

    return activity.type === ActivityTypes.Invoke && activity.name === verifyStateOperationName;
  }

  /**
   * @internal
   */
  private isTokenExchangeRequestInvoke(context: TurnContext): boolean {
    const activity: Activity = context.activity;

    return activity.type === ActivityTypes.Invoke && activity.name === tokenExchangeOperationName;
  }

  /**
   * @internal
   */
  private isTokenExchangeRequest(obj: any): obj is TokenExchangeInvokeRequest {
    return obj.hasOwnProperty("token");
  }
}

/**
 * @internal
 */
interface teamsBotSsoPromptState {
  state: any;
  options: PromptOptions;
  expires: number; // Timestamp of when the prompt will timeout.
}
