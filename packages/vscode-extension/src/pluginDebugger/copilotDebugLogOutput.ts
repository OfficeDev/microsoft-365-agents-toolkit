// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LogLevel } from "@microsoft/teamsfx-api";
import * as vscode from "vscode";
import { ANSIColors } from "../debug/common/debugConstants";
import { defaultExtensionLogPath } from "../globalVariables";

interface Plugin {
  name: string;
  id: string;
  version: string;
}

interface FunctionDescriptor {
  plugin: Plugin;
  functionDisplayName: string;
}

interface FunctionExecution {
  function: FunctionDescriptor;
  executionStatus: {
    requestStatus: number;
    responseStatus: number;
    responseType: number;
  };
  parameters: Record<string, string>;
  requestUri: string;
  requestMethod: string;
  responseContent: string;
  responseContentType: string;
  errorMessage: string;
}

/**
 * @Sample [2021-03-15T03:41:04.961Z] - 0 plugin enabled.
 */
export function logToDebugConsole(logLevel: LogLevel, message: string): void {
  try {
    const dateString = new Date().toJSON();
    const debugConsole = vscode.debug.activeDebugConsole;
    if (logLevel === LogLevel.Info) {
      debugConsole.appendLine(
        ANSIColors.WHITE + `[${dateString}] - ` + ANSIColors.BLUE + `${message}`
      );
    } else if (logLevel === LogLevel.Warning) {
      debugConsole.appendLine(
        ANSIColors.WHITE + `[${dateString}] - ` + ANSIColors.YELLOW + `${message}`
      );
    } else if (logLevel === LogLevel.Error) {
      debugConsole.appendLine(
        ANSIColors.WHITE + `[${dateString}] - ` + ANSIColors.RED + `${message}`
      );
    } else if (logLevel === LogLevel.Debug) {
      debugConsole.appendLine(
        ANSIColors.WHITE + `[${dateString}] - ` + ANSIColors.GREEN + `${message}`
      );
    } else {
      debugConsole.appendLine(ANSIColors.WHITE + `[${dateString}] - ${message}`);
    }
  } catch (e) {}
}

export async function writeCopilotLogToFile(log: string, filePath: string): Promise<void> {
  const fs = require("fs");
  await fs.appendFile(filePath, log + "\n");
}

export class CopilotDebugLog {
  enabledPlugins?: Plugin[];
  matchedFunctionCandidates?: FunctionDescriptor[];
  functionsSelectedForInvocation?: FunctionDescriptor[];
  functionExecutions?: FunctionExecution[];

  constructor(logAsJson: string) {
    let message: this;
    try {
      message = JSON.parse(logAsJson) as this;
    } catch (error) {
      throw new Error(`Error parsing logAsJson: ${(error as Error).message}`);
    }
    this.enabledPlugins = message.enabledPlugins;
    this.matchedFunctionCandidates = message.matchedFunctionCandidates;
    this.functionsSelectedForInvocation = message.functionsSelectedForInvocation;
    this.functionExecutions = message.functionExecutions;

    if (this.functionExecutions) {
      this.functionExecutions.forEach((functionExecution) => {
        try {
          if (functionExecution.requestUri) {
            new URL(functionExecution.requestUri);
          }
        } catch (error) {
          throw new Error(
            `Error creating URL object for requestUri: ${functionExecution.requestUri}`
          );
        }
      });
    }
  }

  write(): void {
    const debugConsole = vscode.debug.activeDebugConsole;
    if (this.enabledPlugins && this.enabledPlugins.length > 0) {
      debugConsole.appendLine("");
      logToDebugConsole(LogLevel.Info, `${this.enabledPlugins.length} enabled plugin(s).`);
      debugConsole.appendLine(ANSIColors.WHITE + "Copilot plugin developer info:");
      debugConsole.appendLine("");
      this.enabledPlugins.forEach((plugin) => {
        debugConsole.appendLine(
          `${ANSIColors.GREEN}(√) ${ANSIColors.WHITE}Enabled plugin: ${ANSIColors.MAGENTA}${plugin.name} ${ANSIColors.GRAY}• version ${plugin.version} • ${plugin.id}`
        );

        if (this.matchedFunctionCandidates && this.matchedFunctionCandidates.length > 0) {
          this.matchedFunctionCandidates.forEach((matchedFunction) => {
            if (matchedFunction.plugin.id === plugin.id) {
              debugConsole.appendLine(
                `${ANSIColors.GREEN}   (√) ${ANSIColors.WHITE}Matched functions: ${ANSIColors.MAGENTA}${matchedFunction.functionDisplayName}`
              );
              this.logFunctionExecutions(debugConsole, matchedFunction);
            }
          });
        }
      });
    } else {
      debugConsole.appendLine("");
      logToDebugConsole(LogLevel.Info, `0 enabled plugin(s).`);
      debugConsole.appendLine(ANSIColors.WHITE + "Copilot plugin developer info:");
      debugConsole.appendLine("");
      this.logNoPlugins(debugConsole);
    }
  }

  private logNoMatchedFunctions(debugConsole: vscode.DebugConsole): void {
    debugConsole.appendLine(
      `${ANSIColors.RED}   (×) Error: ${ANSIColors.WHITE}Matched functions: None`
    );
  }

  private logNoPlugins(debugConsole: vscode.DebugConsole): void {
    debugConsole.appendLine(`${ANSIColors.RED}(×) Error: ${ANSIColors.WHITE}Enabled plugin: None`);
  }

  private logFunctionExecutions(
    debugConsole: vscode.DebugConsole,
    matchedFunction: FunctionDescriptor
  ): void {
    if (this.functionsSelectedForInvocation && this.functionsSelectedForInvocation.length > 0) {
      this.functionsSelectedForInvocation.forEach((selectedFunction) => {
        if (selectedFunction.functionDisplayName === matchedFunction.functionDisplayName) {
          debugConsole.appendLine(
            `${ANSIColors.GREEN}      (√) ${ANSIColors.WHITE}Selected functions for execution: ${ANSIColors.MAGENTA}${selectedFunction.functionDisplayName}`
          );
          this.logExecutionDetails(debugConsole, matchedFunction);
        }
      });
    }
  }

  private logExecutionDetails(
    debugConsole: vscode.DebugConsole,
    matchedFunction: FunctionDescriptor
  ): void {
    const logFileName = `Copilot log ${new Date().toISOString().replace(/-|:|\.\d+Z$/g, "")}.txt`;
    const logFilePath = `${defaultExtensionLogPath}/${logFileName}`;
    if (this.functionExecutions && this.functionExecutions.length > 0) {
      this.functionExecutions.forEach((functionExecution) => {
        if (
          functionExecution.function.functionDisplayName === matchedFunction.functionDisplayName
        ) {
          debugConsole.appendLine(
            `${ANSIColors.GREEN}         (√) ${ANSIColors.WHITE}Function execution details: ${ANSIColors.GREEN}Status ${functionExecution.executionStatus.responseStatus}, ${ANSIColors.WHITE}refer to ${ANSIColors.BLUE}${logFilePath}${ANSIColors.WHITE} for all details.`
          );
          if (functionExecution.errorMessage) {
            debugConsole.appendLine(
              `${ANSIColors.RED}            (×) Error: ${ANSIColors.WHITE}${functionExecution.errorMessage}`
            );
          }
        }
      });
    }
  }

  static prettyPrintJson(jsonText: string): string {
    return JSON.stringify(JSON.parse(jsonText), null, 2);
  }
}
