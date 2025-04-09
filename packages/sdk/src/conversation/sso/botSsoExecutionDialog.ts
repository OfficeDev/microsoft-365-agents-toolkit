// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ComponentDialog,
  WaterfallDialog,
  Dialog,
  DialogSet,
  DialogTurnStatus,
  DialogContext,
} from "@microsoft/agents-hosting-dialogs";
import {
  Activity,
  ActivityTypes,
  Channels,
  AgentStatePropertyAccessor,
  Storage,
  TurnContext,
} from "@microsoft/agents-hosting";
import { tokenExchangeOperationName } from "@microsoft/agents-hosting-teams";
import { CommandMessage, BotSsoExecutionDialogHandler, TriggerPatterns } from "../interface";
import { TeamsBotSsoPrompt, TeamsBotSsoPromptSettings } from "../../bot/teamsBotSsoPrompt";
import { TeamsBotSsoPromptTokenResponse } from "../../bot/teamsBotSsoPromptTokenResponse";
import { formatString } from "../../util/utils";
import { ErrorCode, ErrorMessage, ErrorWithCode } from "../../core/errors";
import { internalLogger } from "../../util/logger";
import { createHash } from "crypto";
import { OnBehalfOfCredentialAuthConfig } from "../../models/configuration";

let DIALOG_NAME = "BotSsoExecutionDialog";
let TEAMS_SSO_PROMPT_ID = "TeamsFxSsoPrompt";
let COMMAND_ROUTE_DIALOG = "CommandRouteDialog";

/**
 * Sso execution dialog, use to handle sso command
 */
export class BotSsoExecutionDialog extends ComponentDialog {
  private dedupStorage: Storage;
  private dedupStorageKeys: string[] = [];

  // Map to store the commandId and triggerPatterns, key: commandId, value: triggerPatterns
  private commandMapping: Map<string, string | RegExp | (string | RegExp)[]> = new Map<
    string,
    string | RegExp | (string | RegExp)[]
  >();

  /**
   * Creates a new instance of the BotSsoExecutionDialog.
   * @param {@link Storage} dedupStorage Helper storage to remove duplicated messages
   * @param {@link TeamsBotSsoPromptSettings} settings The list of scopes for which the token will have access
   * @param {@link OnBehalfOfCredentialAuthConfig} authConfig The authentication configuration.
   * @param {string} initiateLoginEndpoint Login URL for Teams to redirect to.
   * @param {string} dialogName custom dialog name
   */
  constructor(
    dedupStorage: Storage,
    ssoPromptSettings: TeamsBotSsoPromptSettings,
    authConfig: OnBehalfOfCredentialAuthConfig,
    initiateLoginEndpoint: string,
    dialogName?: string
  ) {
    super(dialogName ?? DIALOG_NAME);

    if (dialogName) {
      DIALOG_NAME = dialogName;
      TEAMS_SSO_PROMPT_ID = dialogName + TEAMS_SSO_PROMPT_ID;
      COMMAND_ROUTE_DIALOG = dialogName + COMMAND_ROUTE_DIALOG;
    }

    const ssoDialog = new TeamsBotSsoPrompt(
      authConfig,
      initiateLoginEndpoint,
      TEAMS_SSO_PROMPT_ID,
      ssoPromptSettings
    );
    this.addDialog(ssoDialog);
    this.initialDialogId = COMMAND_ROUTE_DIALOG;
    this.dedupStorage = dedupStorage;
    this.dedupStorageKeys = [];

    const commandRouteDialog = new WaterfallDialog(COMMAND_ROUTE_DIALOG, [
      this.commandRouteStep.bind(this),
    ]);
    this.addDialog(commandRouteDialog);
  }

  /**
   * Add TeamsFxBotSsoCommandHandler instance
   * @param handler {@link BotSsoExecutionDialogHandler} callback function
   * @param triggerPatterns The trigger pattern
   */
  public addCommand(handler: BotSsoExecutionDialogHandler, triggerPatterns: TriggerPatterns): void {
    const commandId = this.getCommandHash(triggerPatterns);
    const dialog = new WaterfallDialog(commandId, [
      this.ssoStep.bind(this),
      this.dedupStep.bind(this),
      async (stepContext: any) => {
        const tokenResponse: TeamsBotSsoPromptTokenResponse = stepContext.result.tokenResponse;
        const context: TurnContext = stepContext.context;
        const message: CommandMessage = stepContext.result.message;

        try {
          if (tokenResponse) {
            await handler(context, tokenResponse, message);
          } else {
            throw new Error(ErrorMessage.FailedToRetrieveSsoToken);
          }
          return await stepContext.endDialog();
        } catch (error: unknown) {
          const errorMsg = formatString(
            ErrorMessage.FailedToProcessSsoHandler,
            (error as Error).message
          );
          internalLogger.error(errorMsg);
          return await stepContext.endDialog(
            new ErrorWithCode(errorMsg, ErrorCode.FailedToProcessSsoHandler)
          );
        }
      },
    ]);

    this.commandMapping.set(commandId, triggerPatterns);
    this.addDialog(dialog);
  }

  private getCommandHash(patterns: TriggerPatterns): string {
    const expressions = Array.isArray(patterns) ? patterns : [patterns];
    const patternStr = expressions.join();
    const patternStrWithoutSpecialChar = patternStr.replace(/[^a-zA-Z0-9]/g, "");
    const hash = createHash("sha256").update(patternStr).digest("hex").toLowerCase();
    return patternStrWithoutSpecialChar + hash;
  }

  /**
   * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
   *
   * @param context The context object for the current turn.
   * @param accessor The instance of StatePropertyAccessor for dialog system.
   */
  public async run(context: TurnContext, accessor: AgentStatePropertyAccessor) {
    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);

    const dialogContext = await dialogSet.createContext(context);
    this.ensureMsTeamsChannel(dialogContext);
    const results = await dialogContext.continueDialog();
    if (results && results.status === DialogTurnStatus.empty) {
      await dialogContext.beginDialog(this.id);
    } else if (
      results &&
      results.status === DialogTurnStatus.complete &&
      results.result instanceof Error
    ) {
      throw results.result;
    }
  }

  private getActivityText(activity: Activity): string {
    let text = activity.text!;
    const removedMentionText = activity.removeRecipientMention();
    if (removedMentionText) {
      text = removedMentionText
        .toLowerCase()
        .replace(/\n|\r\n/g, "")
        .trim();
    }

    return text;
  }

  private async commandRouteStep(stepContext: any) {
    const turnContext = stepContext.context as TurnContext;

    const text = this.getActivityText(turnContext.activity);

    const commandId = this.getMatchesCommandId(text);
    if (commandId) {
      return await stepContext.beginDialog(commandId);
    }

    const errorMsg = formatString(ErrorMessage.CannotFindCommand, turnContext.activity.text!);
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.CannotFindCommand);
  }

  private async ssoStep(stepContext: any) {
    try {
      const turnContext = stepContext.context as TurnContext;

      const text = this.getActivityText(turnContext.activity);
      const message: CommandMessage = {
        text,
      };

      stepContext.options.commandMessage = message;

      return await stepContext.beginDialog(TEAMS_SSO_PROMPT_ID);
    } catch (error: unknown) {
      const errorMsg = formatString(ErrorMessage.FailedToRunSsoStep, (error as Error).message);
      internalLogger.error(errorMsg);
      return await stepContext.endDialog(new ErrorWithCode(errorMsg, ErrorCode.FailedToRunSsoStep));
    }
  }

  private async dedupStep(stepContext: any) {
    const tokenResponse = stepContext.result;
    if (!tokenResponse) {
      internalLogger.error(ErrorMessage.FailedToRetrieveSsoToken);
      return await stepContext.endDialog(
        new ErrorWithCode(ErrorMessage.FailedToRetrieveSsoToken, ErrorCode.FailedToRunSsoStep)
      );
    }

    try {
      // Only dedup after ssoStep to make sure that all Teams client would receive the login request
      if (tokenResponse && (await this.shouldDedup(stepContext.context))) {
        return Dialog.EndOfTurn;
      }
      return await stepContext.next({
        tokenResponse,
        message: stepContext.options.commandMessage,
      });
    } catch (error: unknown) {
      const errorMsg = formatString(ErrorMessage.FailedToRunDedupStep, (error as Error).message);
      internalLogger.error(errorMsg);
      return await stepContext.endDialog(
        new ErrorWithCode(errorMsg, ErrorCode.FailedToRunDedupStep)
      );
    }
  }

  /**
   * Called when the component is ending.
   *
   * @param context Context for the current turn of conversation.
   */
  protected async onEndDialog(context: TurnContext) {
    const conversationId = context.activity.conversation!.id;
    const currentDedupKeys = this.dedupStorageKeys.filter((key) => key.indexOf(conversationId) > 0);
    await this.dedupStorage.delete(currentDedupKeys);
    this.dedupStorageKeys = this.dedupStorageKeys.filter((key) => key.indexOf(conversationId) < 0);
  }

  /**
   * If a user is signed into multiple Teams clients, the Bot might receive a "signin/tokenExchange" from each client.
   * Each token exchange request for a specific user login will have an identical activity.value.Id.
   * Only one of these token exchange requests should be processed by the bot. For a distributed bot in production,
   * this requires a distributed storage to ensure only one token exchange is processed.
   * @param context Context for the current turn of conversation.
   * @returns boolean value indicate whether the message should be removed
   */
  private async shouldDedup(context: TurnContext): Promise<boolean> {
    const storeItem = {
      eTag: (context.activity.value as any).id,
    };

    const key = this.getStorageKey(context);
    const storeItems = { [key]: storeItem };

    try {
      await this.dedupStorage.write(storeItems);
      this.dedupStorageKeys.push(key);
    } catch (err: unknown) {
      if (err instanceof Error && err.message.indexOf("eTag conflict")) {
        return true;
      }
      throw err;
    }
    return false;
  }

  private getStorageKey(context: TurnContext): string {
    if (!context || !context.activity || !context.activity.conversation) {
      throw new Error("Invalid context, can not get storage key!");
    }
    const activity = context.activity;
    const channelId = activity.channelId!;
    const conversationId = activity.conversation!.id;
    if (activity.type !== ActivityTypes.Invoke || activity.name !== tokenExchangeOperationName) {
      throw new Error("TokenExchangeState can only be used with Invokes of signin/tokenExchange.");
    }
    const value = activity.value;
    if (!value || !(value as any).id) {
      throw new Error("Invalid signin/tokenExchange. Missing activity.value.id.");
    }
    return `${channelId}/${conversationId}/${(value as any).id as string}`;
  }

  private matchPattern(pattern: string | RegExp, text: string): boolean | RegExpMatchArray {
    if (text) {
      if (typeof pattern === "string") {
        const regExp = new RegExp(pattern, "i");
        return regExp.test(text);
      }

      if (pattern instanceof RegExp) {
        const matches = text.match(pattern);
        return matches ?? false;
      }
    }

    return false;
  }

  private isPatternMatched(patterns: TriggerPatterns, text: string): boolean {
    const expressions = Array.isArray(patterns) ? patterns : [patterns];

    for (const ex of expressions) {
      const matches = this.matchPattern(ex, text);
      return !!matches;
    }

    return false;
  }

  private getMatchesCommandId(text: string): string | undefined {
    for (const command of this.commandMapping) {
      const pattern: TriggerPatterns = command[1];

      if (this.isPatternMatched(pattern, text)) {
        return command[0];
      }
    }

    return undefined;
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
        "SSO execution dialog"
      );
      internalLogger.error(errorMsg);
      throw new ErrorWithCode(errorMsg, ErrorCode.ChannelNotSupported);
    }
  }
}
