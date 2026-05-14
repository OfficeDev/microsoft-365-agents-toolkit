// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LogLevel, LogProvider, Colors } from "@microsoft/teamsfx-core-next";
import chalk from "chalk";
import { TextType, colorize } from "./output/colorize";

class CLILogProvider implements LogProvider {
  private _logLevel: LogLevel = LogLevel.Info;
  private logFilePath = "";

  getLogLevel(): LogLevel {
    return this._logLevel;
  }

  setLogLevel(level: LogLevel): void {
    this._logLevel = level;
  }

  log(logLevel: LogLevel, message: string): void {
    switch (logLevel) {
      case LogLevel.Debug:
        void this.debug(message);
        break;
      case LogLevel.Verbose:
        void this.verbose(message);
        break;
      case LogLevel.Info:
        void this.info(message);
        break;
      case LogLevel.Warning:
        void this.warning(message);
        break;
      case LogLevel.Error:
        void this.error(message);
        break;
    }
  }

  verbose(message: string): void {
    if (this._logLevel <= LogLevel.Verbose) {
      console.log(colorize(message, TextType.Details));
    }
  }

  debug(message: string): void {
    if (this._logLevel <= LogLevel.Debug) {
      console.log(colorize(message, TextType.Details));
    }
  }

  info(message: string | Array<{ content: string; color: Colors }>): void {
    if (this._logLevel <= LogLevel.Info) {
      if (typeof message === "string") {
        console.log(message);
      } else {
        const formatted = message.map((part) => applyColor(part.content, part.color)).join("");
        console.log(formatted);
      }
    }
  }

  warning(message: string): void {
    if (this._logLevel <= LogLevel.Warning) {
      console.warn(colorize(message, TextType.Warning));
    }
  }

  error(message: string): void {
    if (this._logLevel <= LogLevel.Error) {
      console.error(colorize(message, TextType.Error));
    }
  }

  async logInFile(logLevel: LogLevel, message: string): Promise<void> {
    // File logging will be implemented when log file path is configured
    void logLevel;
    void message;
  }

  getLogFilePath(): string {
    return this.logFilePath;
  }
}

function applyColor(text: string, color: Colors): string {
  if (!process.stdout.isTTY) return text;
  switch (color) {
    case Colors.BRIGHT_GREEN:
      return chalk.greenBright(text);
    case Colors.BRIGHT_RED:
      return chalk.redBright(text);
    case Colors.BRIGHT_YELLOW:
      return chalk.yellowBright(text);
    case Colors.BRIGHT_CYAN:
      return chalk.cyanBright(text);
    case Colors.BRIGHT_MAGENTA:
      return chalk.magentaBright(text);
    case Colors.BRIGHT_WHITE:
      return chalk.whiteBright(text);
    case Colors.WHITE:
      return chalk.white(text);
    default:
      return text;
  }
}

export const logger = new CLILogProvider();
