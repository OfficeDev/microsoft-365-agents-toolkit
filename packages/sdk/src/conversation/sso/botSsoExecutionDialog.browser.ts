// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { formatString } from "../../util/utils";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../../core/errors";
import { BotSsoExecutionDialogHandler, TriggerPatterns } from "../interface";
import { AgentStatePropertyAccessor, TurnContext, Storage } from "@microsoft/agents-hosting";
import { OnBehalfOfCredentialAuthConfig } from "../../models/configuration";
import { TeamsBotSsoPromptSettings } from "../../bot/teamsBotSsoPrompt.browser";

/*
 * Sso execution dialog, use to handle sso command
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export class BotSsoExecutionDialog {
  /**
   * Constructor of BotSsoExecutionDialog
   *
   * @remarks
   * Can Only works in in server side.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  constructor(
    dedupStorage: Storage,
    ssoPromptSettings: TeamsBotSsoPromptSettings,
    authConfig: OnBehalfOfCredentialAuthConfig,
    initiateLoginEndpoint: string,
    dialogName?: string
  );
  constructor(
    dedupStorage: Storage,
    ssoPromptSettings: TeamsBotSsoPromptSettings,
    authConfig: OnBehalfOfCredentialAuthConfig,
    ...args: any
  ) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "BotSsoExecutionDialog"),
      ErrorCode.RuntimeNotSupported
    );
  }

  // eslint-disable-next-line no-secrets/no-secrets
  /**
   * Add TeamsFxBotSsoCommandHandler instance
   * @param handler {@link BotSsoExecutionDialogHandler} callback function
   * @param triggerPatterns The trigger pattern
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  public addCommand(handler: BotSsoExecutionDialogHandler, triggerPatterns: TriggerPatterns): void {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "BotSsoExecutionDialog"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
   *
   * @param context The context object for the current turn.
   * @param accessor The instance of StatePropertyAccessor for dialog system.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  public run(context: TurnContext, accessor: AgentStatePropertyAccessor) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "BotSsoExecutionDialog"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Called when the component is ending.
   *
   * @param context Context for the current turn of conversation.
   */
  protected onEndDialog(context: TurnContext) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "BotSsoExecutionDialog"),
      ErrorCode.RuntimeNotSupported
    );
  }
}
