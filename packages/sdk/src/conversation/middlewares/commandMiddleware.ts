// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Activity, ActivityTypes } from "@microsoft/agents-activity";
import { TurnContext } from "@microsoft/agents-hosting";
import { Middleware } from "@microsoft/agents-hosting/dist/src/middlewareSet";
import { TeamsBotSsoPromptTokenResponse } from "../../bot/teamsBotSsoPromptTokenResponse";
import { ErrorCode, ErrorMessage, ErrorWithCode } from "../../core/errors";
import { internalLogger } from "../../util/logger";
import {
  CommandMessage,
  BotSsoExecutionActivityHandler,
  TeamsFxBotCommandHandler,
  TeamsFxBotSsoCommandHandler,
  TriggerPatterns,
} from "../interface";

/**
 * @internal
 */
export class CommandResponseMiddleware implements Middleware {
  public readonly commandHandlers: TeamsFxBotCommandHandler[] = [];
  public readonly ssoCommandHandlers: TeamsFxBotSsoCommandHandler[] = [];

  public ssoActivityHandler: BotSsoExecutionActivityHandler | undefined;
  public hasSsoCommand: boolean;

  constructor(
    handlers?: TeamsFxBotCommandHandler[],
    ssoHandlers?: TeamsFxBotSsoCommandHandler[],
    activityHandler?: BotSsoExecutionActivityHandler
  ) {
    handlers = handlers ?? [];
    ssoHandlers = ssoHandlers ?? [];
    this.hasSsoCommand = ssoHandlers.length > 0;
    this.ssoActivityHandler = activityHandler;

    if (this.hasSsoCommand && !this.ssoActivityHandler) {
      internalLogger.error(ErrorMessage.SsoActivityHandlerIsNull);
      throw new ErrorWithCode(
        ErrorMessage.SsoActivityHandlerIsNull,
        ErrorCode.SsoActivityHandlerIsUndefined
      );
    }

    this.commandHandlers.push(...handlers);
    for (const ssoHandler of ssoHandlers) {
      this.addSsoCommand(ssoHandler);
    }
  }

  public addSsoCommand(ssoHandler: TeamsFxBotSsoCommandHandler): void {
    this.ssoActivityHandler?.addCommand(
      async (
        context: TurnContext,
        tokenResponse: TeamsBotSsoPromptTokenResponse,
        message: CommandMessage
      ) => {
        const matchResult = this.shouldTrigger(ssoHandler.triggerPatterns, message.text);
        message.matches = Array.isArray(matchResult) ? matchResult : void 0;
        const response = await ssoHandler.handleCommandReceived(context, message, tokenResponse);
        await this.processResponse(context, response);
      },
      ssoHandler.triggerPatterns
    );
    this.ssoCommandHandlers.push(ssoHandler);
    this.hasSsoCommand = true;
  }

  public async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
    if (context.activity.type === ActivityTypes.Message) {
      // Invoke corresponding command handler for the command response
      const commandText = this.getActivityText(context.activity);
      let alreadyProcessed = false;
      for (const handler of this.commandHandlers) {
        const matchResult = this.shouldTrigger(handler.triggerPatterns, commandText);

        // It is important to note that the command bot will stop processing handlers
        // when the first command handler is matched.
        if (!!matchResult) {
          const message: CommandMessage = {
            text: commandText,
          };
          message.matches = Array.isArray(matchResult) ? matchResult : void 0;
          const response = await handler.handleCommandReceived(context, message);

          await this.processResponse(context, response);
          alreadyProcessed = true;
          break;
        }
      }

      if (!alreadyProcessed) {
        for (const handler of this.ssoCommandHandlers) {
          const matchResult = this.shouldTrigger(handler.triggerPatterns, commandText);
          if (!!matchResult) {
            await this.ssoActivityHandler?.run(context);
            break;
          }
        }
      }
    } else {
      if (this.hasSsoCommand) {
        await this.ssoActivityHandler?.run(context);
      }
    }
    await next();
  }

  private async processResponse(context: TurnContext, response: string | void | Partial<Activity>) {
    if (typeof response === "string") {
      await context.sendActivity(response);
    } else {
      const replyActivity = response as Partial<Activity>;
      if (replyActivity) {
        await context.sendActivity(replyActivity as Activity);
      }
    }
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

  private shouldTrigger(patterns: TriggerPatterns, text: string): RegExpMatchArray | boolean {
    const expressions = Array.isArray(patterns) ? patterns : [patterns];

    for (const ex of expressions) {
      const arg = this.matchPattern(ex, text);
      if (arg) return arg;
    }

    return false;
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
}
