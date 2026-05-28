// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { find } from "lodash";
import * as cp from "child_process";
import * as vscode from "vscode";
import * as globalVariables from "../../globalVariables";
import { FxError, Result, Void, ok } from "@microsoft/teamsfx-api";
// eslint-disable-next-line import/no-cycle
import { BaseTaskTerminal, ControlCodes } from "./baseTaskTerminal";
import { OfficeManifestType, fetchManifestList } from "@microsoft/teamsfx-core";
import { localize } from "../../utils/localizeUtils";

export enum TriggerCmdType {
  triggerInstall = "trigger install dependencies",
  triggerValidate = "trigger validate",
  triggerStopDebug = "trigger stop debug",
  triggerGenerateGUID = "generate manifest GUID",
}

enum ProcessStatus {
  notStarted,
  running,
  completed,
}

export class OfficeDevTerminal extends BaseTaskTerminal {
  private status = ProcessStatus.notStarted;

  constructor() {
    super();
  }

  do(): Promise<Result<Void, FxError>> {
    return Promise.resolve(ok(Void));
  }

  async open() {
    await this.do();
  }

  close(): void {
    this.stop().catch((error) => {
      this.writeEmitter.fire(`${error?.message as string}\r\n`);
    });
  }

  handleInput(data: string): void {
    if (data.includes(ControlCodes.CtrlC)) {
      this.stop().catch((error) => {
        this.writeEmitter.fire(`${error?.message as string}\r\n`);
      });
    } else if (data.startsWith(TriggerCmdType.triggerInstall)) {
      if (this.status != ProcessStatus.running) {
        this.writeEmitter.fire(
          `\r\n${this.color(
            localize("teamstoolkit.officeAddIn.terminal.installDependency"),
            "yellow"
          )}\r\n`
        );
        this.installDependencies();
        this.status = ProcessStatus.running;
      }
    } else if (data.startsWith(TriggerCmdType.triggerValidate)) {
      if (this.status != ProcessStatus.running) {
        this.writeEmitter.fire(
          `\r\n${this.color(
            localize("teamstoolkit.officeAddIn.terminal.validateManifest"),
            "yellow"
          )}\r\n`
        );
        this.runValidate();
        this.status = ProcessStatus.running;
      }
    } else if (data.startsWith(TriggerCmdType.triggerStopDebug)) {
      if (this.status != ProcessStatus.running) {
        this.writeEmitter.fire(
          `\r\n${this.color(
            localize("teamstoolkit.officeAddIn.terminal.stopDebugging"),
            "yellow"
          )}\r\n`
        );
        this.stopDebug();
        this.status = ProcessStatus.running;
      }
    } else if (data.startsWith(TriggerCmdType.triggerGenerateGUID)) {
      if (this.status != ProcessStatus.running) {
        this.writeEmitter.fire(
          `\r\n${this.color(
            localize("teamstoolkit.officeAddIn.terminal.generateManifestGUID"),
            "yellow"
          )}\r\n`
        );
        this.generateManifestGUID();
        this.status = ProcessStatus.running;
      }
    } else if (this.status == ProcessStatus.completed) {
      this.closeEmitter.fire(0);
    }
  }

  public startChildProcess(cmd: string, args: readonly string[]) {
    const cmdStr = cmd + " " + args.join(" ");
    this.writeEmitter.fire(`${cmdStr}\r\n`);
    const options: cp.SpawnOptions = {
      cwd: globalVariables.workspaceUri?.fsPath ?? "" + "/src",
      shell: true,
      detached: false,
    };

    const childProc = cp.spawn(cmd, args, options);
    childProc.stdout?.setEncoding("utf-8");
    childProc.stdout?.on("data", (data: string | Buffer) => {
      const line = data.toString().replace(/\n/g, "\r\n");
      this.writeEmitter.fire(line);
    });

    childProc.stderr?.setEncoding("utf-8");
    childProc.stderr?.on("data", (data: string | Buffer) => {
      const line = data.toString().replace(/\n/g, "\r\n");
      this.writeEmitter.fire(line);
    });

    childProc.on("exit", () => {
      this.writeEmitter.fire(localize("teamstoolkit.officeAddIn.terminal.terminate") + "\r\n");
      this.status = ProcessStatus.completed;
    });
  }

  public stopDebug() {
    const manifestFileName = this.getManifest();
    if (manifestFileName) {
      const cmd = "npx";
      const args = ["office-addin-debugging", "stop", manifestFileName];
      this.startChildProcess(cmd, args);
    }
  }

  public generateManifestGUID() {
    const manifestFileName = this.getManifest();
    if (manifestFileName) {
      const cmd = "npx";
      const args = ["--yes", "office-addin-manifest", "modify", manifestFileName, "--guid"];
      this.startChildProcess(cmd, args);
    }
  }

  public runValidate() {
    const manifestFileName = this.getManifest();
    if (manifestFileName) {
      const cmd = "npx";
      const args = ["--yes", "office-addin-manifest", "validate", manifestFileName];
      this.startChildProcess(cmd, args);
    }
  }

  public installDependencies() {
    const cmd = "npm";
    const args = ["install", "--color=always"];
    this.startChildProcess(cmd, args);
  }

  private getManifest(): string | undefined {
    const workspacePath = globalVariables.workspaceUri?.fsPath;
    const manifestList = fetchManifestList(workspacePath, OfficeManifestType.XmlAddIn);
    if (!manifestList || manifestList.length == 0) {
      this.writeEmitter.fire(
        this.color(`${localize("teamstoolkit.officeAddIn.terminal.manifest.notfound")}\r\n`, "red")
      );
      return undefined;
    }

    return manifestList[0];
  }

  private color(msg: string, color: string) {
    switch (color) {
      case "red":
        return "\x1b[31m" + msg + "\x1b[0m";
      case "green":
        return "\x1b[32m" + msg + "\x1b[0m";
      case "yellow":
        return "\x1b[33m" + msg + "\x1b[0m";
      default:
        return msg;
    }
  }

  public static getTerminalTitle(triggerCmd: TriggerCmdType): string | undefined {
    switch (triggerCmd) {
      case TriggerCmdType.triggerInstall:
        return localize("teamstoolkit.commandsTreeViewProvider.checkAndInstallDependenciesTitle");
      case TriggerCmdType.triggerGenerateGUID:
        return localize("teamstoolkit.codeLens.generateManifestGUID");
      case TriggerCmdType.triggerStopDebug:
        return localize("teamstoolkit.commandsTreeViewProvider.officeAddIn.stopDebugTitle");
      case TriggerCmdType.triggerValidate:
        return localize("teamstoolkit.commandsTreeViewProvider.validateManifestTitle");
      default:
        return undefined;
    }
  }

  public static getInstance(triggerCmd: TriggerCmdType): vscode.Terminal {
    let terminal: vscode.Terminal | undefined;
    const terminalTitle = OfficeDevTerminal.getTerminalTitle(triggerCmd);
    if (
      vscode.window.terminals.length === 0 ||
      (terminal = find(vscode.window.terminals, (value) => value.name === terminalTitle)) ===
        undefined
    ) {
      terminal = vscode.window.createTerminal({
        name: terminalTitle || "officeAddInDev task",
        pty: new OfficeDevTerminal(),
      });
    }
    return terminal;
  }
}
