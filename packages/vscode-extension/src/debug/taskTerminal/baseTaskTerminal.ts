// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */
import { performance } from "perf_hooks";
import * as util from "util";
import { v4 as uuidv4 } from "uuid";
import * as vscode from "vscode";
import { FxError, Result, SystemError, UserError, Void } from "@microsoft/teamsfx-api";
import { assembleError, Correlator } from "@microsoft/teamsfx-core";
import { ExtensionErrors, ExtensionSource } from "../../error/error";
import * as globalVariables from "../../globalVariables";
import { showError } from "../../error/common";
import { TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { getDefaultString, localize } from "../../utils/localizeUtils";
import { sendDebugAllEvent } from "../localTelemetryReporter";
import { DebugNoSessionId } from "../common/debugConstants";
import { getLocalDebugSession, endLocalDebugSession } from "../common/localDebugSession";

export const ControlCodes = {
  CtrlC: "\u0003",
};

export abstract class BaseTaskTerminal implements vscode.Pseudoterminal {
  protected writeEmitter = new vscode.EventEmitter<string>();
  onDidWrite: vscode.Event<string> = this.writeEmitter.event;
  protected closeEmitter = new vscode.EventEmitter<number>();
  onDidClose?: vscode.Event<number> = this.closeEmitter.event;
  protected readonly taskTerminalId: string;
  protected startTime: number | undefined;

  constructor(private taskDefinition?: vscode.TaskDefinition) {
    this.taskTerminalId = uuidv4();
  }

  open(): void {
    this.startTime = performance.now();
    this.do()
      .then((res) => this.stop(res.isErr() ? res.error : undefined))
      .catch((error) => this.stop(error));
  }

  close(): void {
    this.stop().catch((error) => {
      this.writeEmitter.fire(`${error?.message as string}\r\n`);
    });
  }

  handleInput(data: string): void {
    if (data.includes(ControlCodes.CtrlC)) {
      this.stop(
        new UserError(
          ExtensionSource,
          ExtensionErrors.TaskCancelError,
          getDefaultString("teamstoolkit.localDebug.taskCancelError"),
          localize("teamstoolkit.localDebug.taskCancelError")
        )
      ).catch((error) => {
        this.writeEmitter.fire(`${error?.message as string}\r\n`);
      });
    }
  }

  protected async stop(error?: any, outputError = true): Promise<void> {
    if (error) {
      const fxError: UserError | SystemError = assembleError(error);

      // TODO: add color
      this.writeEmitter.fire(`${fxError.message}\r\n`);

      if (outputError) {
        showError(fxError).catch(() => {
          // ignore
        });
      }
      this.closeEmitter.fire(1);

      await Correlator.runWithId(getLocalDebugSession().id, () =>
        sendDebugAllEvent(fxError, { [TelemetryProperty.DebugIsTransparentTask]: "true" })
      );
      if (getLocalDebugSession().id !== DebugNoSessionId) {
        endLocalDebugSession();
      }
    }
    this.closeEmitter.fire(0);
  }

  protected abstract do(): Promise<Result<Void, FxError>>;

  protected getDurationInSeconds(): number | undefined {
    if (!this.startTime) {
      return undefined;
    }
    return (performance.now() - this.startTime) / 1000;
  }

  public static resolveTeamsFxVariables(str: string): string {
    // Background task cannot resolve variables in VSC.
    // Here Teams Toolkit resolve the workspaceFolder.
    str = str.replace("${workspaceFolder}", globalVariables.workspaceUri?.fsPath ?? "");
    return str;
  }

  public static taskDefinitionError(argName: string): UserError {
    return new UserError(
      ExtensionSource,
      ExtensionErrors.TaskDefinitionError,
      util.format(getDefaultString("teamstoolkit.localDebug.taskDefinitionError"), argName),
      util.format(localize("teamstoolkit.localDebug.taskDefinitionError"), argName)
    );
  }
}
