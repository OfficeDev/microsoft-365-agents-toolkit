// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

import { FxError, Result, Stage, ok } from "@microsoft/teamsfx-api";
import { Correlator, TaskCommand } from "@microsoft/teamsfx-core";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import { DebugNoSessionId, TeamsFxTaskType } from "./common/debugConstants";
import { getLocalDebugSessionId } from "./common/localDebugSession";
import { localTelemetryReporter } from "./localTelemetryReporter";
import { DevTunnelTaskTerminal } from "./taskTerminal/devTunnelTaskTerminal";
import { LaunchDesktopClientTerminal } from "./taskTerminal/launchDesktopClientTerminal";
import { LaunchTeamsClientTerminal } from "./taskTerminal/launchTeamsClientTerminal";
import { LifecycleTaskTerminal } from "./taskTerminal/lifecycleTaskTerminal";
import { PrerequisiteTaskTerminal } from "./taskTerminal/prerequisiteTaskTerminal";

const customTasks = Object.freeze({
  [TaskCommand.checkPrerequisites]: {
    createTerminal: (d: vscode.TaskDefinition) => Promise.resolve(new PrerequisiteTaskTerminal(d)),
    presentationReveal: vscode.TaskRevealKind.Never,
    presentationEcho: false,
    presentationshowReuseMessage: false,
  },
  [TaskCommand.startLocalTunnel]: {
    createTerminal: (d: vscode.TaskDefinition) => Promise.resolve(DevTunnelTaskTerminal.create(d)),
    presentationReveal: vscode.TaskRevealKind.Silent,
    presentationEcho: true,
    presentationshowReuseMessage: true,
  },
  [TaskCommand.launchWebClient]: {
    createTerminal: (d: vscode.TaskDefinition) => Promise.resolve(new LaunchTeamsClientTerminal(d)),
    presentationReveal: vscode.TaskRevealKind.Never,
    presentationEcho: false,
    presentationshowReuseMessage: false,
  },
  [TaskCommand.provision]: {
    createTerminal: (d: vscode.TaskDefinition) =>
      Promise.resolve(new LifecycleTaskTerminal(d, Stage.provision)),
    presentationReveal: vscode.TaskRevealKind.Never,
    presentationEcho: false,
    presentationshowReuseMessage: false,
  },
  [TaskCommand.deploy]: {
    createTerminal: (d: vscode.TaskDefinition) =>
      Promise.resolve(new LifecycleTaskTerminal(d, Stage.deploy)),
    presentationReveal: vscode.TaskRevealKind.Never,
    presentationEcho: false,
    presentationshowReuseMessage: false,
  },
  [TaskCommand.installApp]: {
    createTerminal: (d: vscode.TaskDefinition) =>
      Promise.resolve(new LifecycleTaskTerminal(d, Stage.installApp)),
    presentationReveal: vscode.TaskRevealKind.Never,
    presentationEcho: false,
    presentationshowReuseMessage: false,
  },
  [TaskCommand.launchDesktopClient]: {
    createTerminal: (d: vscode.TaskDefinition) =>
      Promise.resolve(new LaunchDesktopClientTerminal(d)),
    presentationReveal: vscode.TaskRevealKind.Silent,
    presentationEcho: true,
    presentationshowReuseMessage: true,
  },
});

export class TeamsfxTaskProvider implements vscode.TaskProvider {
  // eslint-disable-next-line @typescript-eslint/require-await
  public async provideTasks(
    token?: vscode.CancellationToken | undefined
  ): Promise<vscode.Task[] | undefined> {
    return undefined;
  }

  public async resolveTask(
    task: vscode.Task,
    token?: vscode.CancellationToken | undefined
  ): Promise<vscode.Task | undefined> {
    return Correlator.runWithId(
      getLocalDebugSessionId(),
      async (): Promise<vscode.Task | undefined> => {
        let resolvedTask: vscode.Task | undefined = undefined;
        if (getLocalDebugSessionId() === DebugNoSessionId) {
          resolvedTask = this._resolveTask(task, token);
        } else {
          // Only send telemetry within a local debug session.
          await localTelemetryReporter.runWithTelemetry(
            TelemetryEvent.DebugTaskProvider,
            () =>
              new Promise<Result<vscode.Task | undefined, FxError>>((resolve) => {
                resolvedTask = this._resolveTask(task, token);
                resolve(ok(resolvedTask));
              })
          );
        }
        return resolvedTask;
      }
    );
  }

  private _resolveTask(
    task: vscode.Task,
    token?: vscode.CancellationToken | undefined
  ): vscode.Task | undefined {
    if (task.definition.type !== TeamsFxTaskType || !task.definition.command) {
      return undefined;
    }

    const customTask = Object.entries(customTasks).find(
      ([k]) => k === task.definition.command
    )?.[1];
    if (!customTask) {
      return undefined;
    }

    const newTask = new vscode.Task(
      task.definition,
      vscode.TaskScope.Workspace,
      task.name,
      TeamsFxTaskType,
      new vscode.CustomExecution(customTask.createTerminal)
    );

    newTask.presentationOptions.reveal = customTask.presentationReveal;
    newTask.presentationOptions.echo = customTask.presentationEcho;
    newTask.presentationOptions.showReuseMessage = customTask.presentationshowReuseMessage;
    return newTask;
  }
}
