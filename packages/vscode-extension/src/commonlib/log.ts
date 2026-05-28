// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { LogLevel, LogProvider, Colors } from "@microsoft/teamsfx-api";
import * as vscode from "vscode";
import fs from "fs-extra";
import { defaultExtensionLogPath } from "../globalVariables";
import { SummaryConstant } from "@microsoft/teamsfx-core";

const outputChannelDisplayName = "Teams Toolkit";

export class VsCodeLogProvider implements LogProvider {
  logLevel: LogLevel = LogLevel.Info;
  outputChannel: vscode.OutputChannel;
  logFileName: string;

  private static instance: VsCodeLogProvider;

  private constructor() {
    this.outputChannel = vscode.window.createOutputChannel(
      outputChannelDisplayName,
      // Align with languages:id in package.json for colorized output.
      "teamsfx-toolkit-output"
    );
    // E.g. "Wed Oct 05 2011 22:48:00 GMT+0800 (China Standard Time)"
    // > "20230328T070957.log"
    this.logFileName = `${new Date().toISOString().replace(/-|:|\.\d+Z$/g, "")}.log`;
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): VsCodeLogProvider {
    if (!VsCodeLogProvider.instance) {
      VsCodeLogProvider.instance = new VsCodeLogProvider();
    }

    return VsCodeLogProvider.instance;
  }

  /**
   * Get log file path
   */
  getLogFilePath(): string {
    return `${defaultExtensionLogPath}/${this.logFileName}`;
  }

  verbose(message: string): void {
    return this.log(LogLevel.Verbose, message);
  }

  debug(message: string): void {
    return this.log(LogLevel.Debug, message);
  }

  info(message: Array<{ content: string; color: Colors }>): void;

  info(message: string): void;

  info(message: string | Array<{ content: string; color: Colors }>): void {
    // VSCode output channel is not TTY, does not support ANSI color
    if (message instanceof Array) {
      message = message.map((x) => x.content).join("");
    }
    this.log(LogLevel.Info, message);
  }

  warning(message: string): void {
    return this.log(LogLevel.Warning, message);
  }

  error(message: string): void {
    return this.log(LogLevel.Error, message);
  }

  /**
   * @Sample [2021-03-15T03:41:04.961Z] [Info] - [Extension] Initialize successfully.
   */
  log(logLevel: LogLevel, message: string): void {
    try {
      if (logLevel < this.logLevel) return;
      if (logLevel >= LogLevel.Warning) this.outputChannel.show();
      const dateString = new Date().toJSON();
      const formattedMessage = `[${dateString}] [${LogLevel[logLevel]}] - ${message}`;
      this.outputChannel.appendLine(formattedMessage);
    } catch (e) {}
  }

  /**
   * @Sample (×) Error: Lifecycle stage deploy failed.
   * @Sample (√) Done: devTool/install was executed successfully.
   */
  semLog(
    messages:
      | Array<{ content: string; status?: SummaryConstant }>
      | { content: string; status?: SummaryConstant }
  ): void {
    try {
      this.outputChannel.show();
      const data: Array<{ content: string; status?: SummaryConstant }> = [];
      if (Array.isArray(messages)) {
        data.push(...messages);
      } else {
        data.push(messages);
      }

      data.forEach((v) => {
        if (v.status) {
          this.outputChannel.appendLine(`${v.status} ${v.content}`);
        } else {
          this.outputChannel.appendLine(v.content);
        }
      });
    } catch (e) {}
    return;
  }

  async logInFile(logLevel: LogLevel, message: string): Promise<void> {
    if (logLevel === LogLevel.Info) {
      const dateString = new Date().toJSON();
      const formattedMessage = `[${dateString}] [${LogLevel[logLevel]}] - ${message}`;
      await fs.appendFile(this.getLogFilePath(), formattedMessage + "\n");
    }
  }
}

export default VsCodeLogProvider.getInstance();
