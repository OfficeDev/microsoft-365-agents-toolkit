// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/require-await */
import { TurnContext } from "@microsoft/agents-hosting";
import { SigninStateVerificationQuery } from "@microsoft/agents-hosting-teams";

import { BotSsoConfig, BotSsoExecutionDialogHandler, TriggerPatterns } from "../interface";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../../core/errors";
import { formatString } from "../../util/utils";

/**
 * Default sso execution activity handler
 */
export class DefaultBotSsoExecutionActivityHandler {
  // eslint-disable-next-line no-secrets/no-secrets
  /**
   * Creates a new instance of the DefaultBotSsoExecutionActivityHandler.
   * @param ssoConfig configuration for sso command bot
   */
  constructor(ssoConfig: BotSsoConfig | undefined) {
    throw new ErrorWithCode(
      formatString(
        ErrorMessage.BrowserRuntimeNotSupported,
        "DefaultBotSsoExecutionActivityHandler"
      ),
      ErrorCode.RuntimeNotSupported
    );
  }

  // eslint-disable-next-line no-secrets/no-secrets
  /**
   * Add TeamsFxBotSsoCommandHandler instance to sso execution dialog
   * @param handler {@link BotSsoExecutionDialogHandler} callback function
   * @param triggerPatterns The trigger pattern
   */
  addCommand(handler: BotSsoExecutionDialogHandler, triggerPatterns: TriggerPatterns): void {
    throw new ErrorWithCode(
      formatString(
        ErrorMessage.BrowserRuntimeNotSupported,
        "DefaultBotSsoExecutionActivityHandler"
      ),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Called to initiate the event emission process.
   * @param context The context object for the current turn.
   */
  async run(context: TurnContext) {
    throw new ErrorWithCode(
      formatString(
        ErrorMessage.BrowserRuntimeNotSupported,
        "DefaultBotSsoExecutionActivityHandler"
      ),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Receives invoke activities with Activity name of 'signin/verifyState'.
   * @param context A context object for this turn.
   * @param query Signin state (part of signin action auth flow) verification invoke query.
   * @returns A promise that represents the work queued.
   */
  async handleTeamsSigninVerifyState(context: TurnContext, query: SigninStateVerificationQuery) {
    throw new ErrorWithCode(
      formatString(
        ErrorMessage.BrowserRuntimeNotSupported,
        "DefaultBotSsoExecutionActivityHandler"
      ),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Receives invoke activities with Activity name of 'signin/tokenExchange'
   * @param context A context object for this turn.
   * @param query Signin state (part of signin action auth flow) verification invoke query
   * @returns A promise that represents the work queued.
   */
  async handleTeamsSigninTokenExchange(context: TurnContext, query: SigninStateVerificationQuery) {
    throw new ErrorWithCode(
      formatString(
        ErrorMessage.BrowserRuntimeNotSupported,
        "DefaultBotSsoExecutionActivityHandler"
      ),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Handle signin invoke activity type.
   *
   * @param context The context object for the current turn.
   *
   * @remarks
   * Override this method to support channel-specific behavior across multiple channels.
   */
  async onSignInInvoke(context: TurnContext) {
    throw new ErrorWithCode(
      formatString(
        ErrorMessage.BrowserRuntimeNotSupported,
        "DefaultBotSsoExecutionActivityHandler"
      ),
      ErrorCode.RuntimeNotSupported
    );
  }
}
