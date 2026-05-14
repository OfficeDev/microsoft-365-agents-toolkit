// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { cliTelemetry } from "./telemetry";
import { colorize, TextType } from "./output";
import { FxError, UserError, registerBuiltinDrivers } from "@microsoft/teamsfx-core-next";
import type { AtkContext, PostAction } from "@microsoft/teamsfx-core-next";
import { createCliContext } from "./context";

export type CommandHandler = (opts: Record<string, unknown>, cmd: Command) => Promise<void>;

export type ContextCommandHandler = (
  ctx: AtkContext,
  opts: Record<string, unknown>,
  cmd: Command
) => Promise<void>;

/**
 * Wrap a command handler with telemetry, error handling, and timing.
 * This is the thin replacement for the old middleware-based engine.
 */
export function wrapHandler(
  commandName: string,
  handler: CommandHandler
): (...args: unknown[]) => void {
  return async (...args: unknown[]) => {
    const cmd = args[args.length - 1] as Command;
    const opts = cmd.opts();

    const startTime = Date.now();
    cliTelemetry.sendEvent(`${commandName}-start`, { command: commandName });

    try {
      await handler(opts, cmd);

      const duration = Date.now() - startTime;
      cliTelemetry.sendEvent(`${commandName}-end`, { command: commandName }, { duration });
    } catch (error: unknown) {
      const duration = Date.now() - startTime;

      if (isFxError(error)) {
        console.error(colorize(`Error: ${error.message}`, TextType.Error));
        if (error instanceof UserError && error.helpLink) {
          console.error(`  Help: ${colorize(error.helpLink, TextType.Hyperlink)}`);
        }
        cliTelemetry.sendErrorEvent(
          `${commandName}-error`,
          error,
          { command: commandName },
          { duration }
        );
      } else if (error instanceof Error) {
        console.error(colorize(`Error: ${error.message}`, TextType.Error));
        cliTelemetry.sendErrorEvent(
          `${commandName}-error`,
          error,
          { command: commandName },
          { duration }
        );
      } else {
        console.error(colorize(`Error: ${String(error)}`, TextType.Error));
      }

      process.exitCode = 1;
    } finally {
      await cliTelemetry.flush();
    }
  };
}

function isFxError(error: unknown): error is FxError {
  return (
    typeof error === "object" &&
    error !== null &&
    "source" in error &&
    "name" in error &&
    "message" in error
  );
}

/**
 * Wrap a command handler that receives an AtkContext.
 * Creates context from Commander options (--folder → projectPath), then delegates.
 */
export function wrapHandlerWithContext(
  commandName: string,
  handler: ContextCommandHandler
): (...args: unknown[]) => void {
  return wrapHandler(commandName, async (opts, cmd) => {
    // Lazy-register drivers & templates — only pay the cost when a real
    // command (not --help) executes.
    registerBuiltinDrivers();
    const projectPath = (opts.folder as string) ?? (opts.projectFolder as string) ?? process.cwd();
    const ctx = createCliContext(projectPath);
    await handler(ctx, opts, cmd);
  });
}

/**
 * Render post-execution actions (messages, URLs) to the console.
 */
export function renderPostActions(actions: PostAction[]): void {
  for (const action of actions) {
    if (action.type === "showMessage") {
      console.log(colorize(action.message, TextType.Success));
    } else if (action.type === "openUrl" && action.url) {
      console.log(`${action.message}: ${colorize(action.url, TextType.Hyperlink)}`);
    }
  }
}
