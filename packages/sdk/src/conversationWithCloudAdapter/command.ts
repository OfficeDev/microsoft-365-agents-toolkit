// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CloudAdapter } from "@microsoft/agents-hosting";
import { ErrorCode, ErrorMessage, ErrorWithCode } from "../core/errors";
import { internalLogger } from "../util/logger";
import {
  CommandOptions,
  BotSsoConfig,
  BotSsoExecutionActivityHandler,
  TeamsFxBotCommandHandler,
  TeamsFxBotSsoCommandHandler,
} from "../conversation/interface";
import { CommandResponseMiddleware } from "../conversation/middlewares/commandMiddleware";

/**
 * A command bot for receiving commands and sending responses in Teams.
 *
 * @remarks
 * Ensure each command should ONLY be registered with the command once, otherwise it'll cause unexpected behavior if you register the same command more than once.
 */
export class CommandBot {
  private readonly adapter: CloudAdapter;
  private readonly middleware: CommandResponseMiddleware;
  private readonly ssoConfig: BotSsoConfig | undefined;

  // eslint-disable-next-line no-secrets/no-secrets
  /**
   * Create a new instance of the `CommandBot`.
   *
   * @param adapter - The bound `CloudAdapter`.
   * @param options - The initialize options
   * @param ssoCommandActivityHandler - SSO execution activity handler.
   * @param ssoConfig - SSO configuration for Bot SSO.
   */
  constructor(
    adapter: CloudAdapter,
    options?: CommandOptions,
    ssoCommandActivityHandler?: BotSsoExecutionActivityHandler,
    ssoConfig?: BotSsoConfig
  ) {
    this.ssoConfig = ssoConfig;

    this.middleware = new CommandResponseMiddleware(
      options?.commands,
      options?.ssoCommands,
      ssoCommandActivityHandler
    );
    this.adapter = adapter.use(this.middleware);
  }

  /**
   * Register a command into the command bot.
   *
   * @param command - The command to be registered.
   */
  public registerCommand(command: TeamsFxBotCommandHandler): void {
    if (command) {
      this.middleware.commandHandlers.push(command);
    }
  }

  /**
   * Register commands into the command bot.
   *
   * @param commands - The commands to be registered.
   */
  public registerCommands(commands: TeamsFxBotCommandHandler[]): void {
    if (commands) {
      this.middleware.commandHandlers.push(...commands);
    }
  }

  /**
   * Register a sso command into the command bot.
   *
   * @param ssoCommand - The sso command to be registered.
   */
  public registerSsoCommand(ssoCommand: TeamsFxBotSsoCommandHandler): void {
    this.validateSsoActivityHandler();
    this.middleware.addSsoCommand(ssoCommand);
  }

  /**
   * Register sso commands into the command bot.
   *
   * @param ssoCommands - The sso commands to be registered.
   */
  public registerSsoCommands(ssoCommands: TeamsFxBotSsoCommandHandler[]): void {
    if (ssoCommands.length > 0) {
      this.validateSsoActivityHandler();
      for (const ssoCommand of ssoCommands) {
        this.middleware.addSsoCommand(ssoCommand);
      }
    }
  }

  private validateSsoActivityHandler() {
    if (!this.middleware.ssoActivityHandler) {
      internalLogger.error(ErrorMessage.SsoActivityHandlerIsNull);
      throw new ErrorWithCode(
        ErrorMessage.SsoActivityHandlerIsNull,
        ErrorCode.SsoActivityHandlerIsUndefined
      );
    }
  }
}
