// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CloudAdapter } from "@microsoft/agents-hosting";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";
import { TeamsFxBotCommandHandler, TeamsFxBotSsoCommandHandler } from "../conversation/interface";
import { CommandResponseMiddleware } from "../conversation/middlewares/commandMiddleware";

/**
 * A command bot for receiving commands and sending responses in Teams.
 *
 * @remarks
 * Only work on server side.
 */
export class CommandBot {
  private readonly adapter: CloudAdapter;
  private readonly middleware: CommandResponseMiddleware;

  /**
   * Create a new instance of the `CommandBot`.
   *
   * @param adapter - The bound `CloudAdapter`.
   * @param commands - The commands to be registered with the command bot. Each command should implement the interface {@link TeamsFxBotCommandHandler} so that it can be correctly handled by this command bot.
   */
  constructor(adapter: CloudAdapter, commands?: TeamsFxBotCommandHandler[]) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CommandBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Register a command into the command bot.
   *
   * @param command - The command to be registered.
   *
   * @remarks
   * Only work on server side.
   */
  public registerCommand(command: TeamsFxBotCommandHandler): void {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CommandBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Register commands into the command bot.
   *
   * @param commands - The commands to be registered.
   *
   * @remarks
   * Only work on server side.
   */
  public registerCommands(commands: TeamsFxBotCommandHandler[]): void {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CommandBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Register a sso command into the command bot.
   *
   * @param ssoCommand - The sso command to be registered.
   */
  public registerSsoCommand(ssoCommand: TeamsFxBotSsoCommandHandler): void {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CommandBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Register sso commands into the command bot.
   *
   * @param ssoCommands - The sso commands to be registered.
   */
  public registerSsoCommands(ssoCommands: TeamsFxBotSsoCommandHandler[]): void {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CommandBot"),
      ErrorCode.RuntimeNotSupported
    );
  }
}
