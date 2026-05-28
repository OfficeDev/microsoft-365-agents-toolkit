// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import { Correlator, isValidOfficeAddInProject } from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import { CommandKey } from "../constants";
import * as globalVariables from "../globalVariables";
import { updateProjectStatus } from "../utils/projectStatusUtils";
import {
  checkAndSkipDebugging,
  endLocalDebugSession,
  getLocalDebugSessionId,
  startLocalDebugSession,
} from "./common/localDebugSession";
import { DebugSessionExists } from "./common/debugConstants";

export const allRunningOfficeTasks: Map<string, number> = new Map<string, number>();
export const allRunningDebugSessions: Set<string> = new Set<string>();

export const OfficeTaskName = [
  "Excel Desktop (Edge Chromium)",
  "Excel Desktop (Edge Legacy)",
  "Debug: Excel Desktop",
  "PowerPoint Desktop (Edge Chromium)",
  "PowerPoint Desktop (Edge Legacy)",
  "Debug: PowerPoint Desktop",
  "Word Desktop (Edge Chromium)",
  "Word Desktop (Edge Legacy)",
  "Debug: Word Desktop",
];

const trackedTasks = new Set<string>();

function getTaskKey(task: vscode.Task): string {
  if (task === undefined) {
    return "";
  }

  // "source|name|scope"
  const scope = (task.scope as vscode.WorkspaceFolder)?.uri?.toString() || task.scope?.toString();
  return `${task.source}|${task.name}|${scope ?? ""}`;
}

function isOfficeTask(task: vscode.Task): boolean {
  if (task) {
    const taskName = task.name;
    if (Object.values(OfficeTaskName).includes(taskName)) {
      return true;
    }
  }
  return false;
}

function isDebugPreLaunchTask(task: vscode.Task): boolean {
  if (task) {
    // Debug: PowerPoint Desktop
    // Debug: Word Desktop
    // Debug: Excel Desktop
    if (task.execution && <vscode.ShellExecution>task.execution) {
      const execution = <vscode.ShellExecution>task.execution;
      const commandLine =
        execution.commandLine ||
        `${typeof execution.command === "string" ? execution.command : execution.command.value} ${(
          execution.args || []
        ).join(" ")}`;
      if (/npm[\s]+run[\s]+start:desktop -- --app (word|excel|powerpoint)/i.test(commandLine)) {
        return true;
      }
    }
  }

  return false;
}

function onDidStartTaskHandler(event: vscode.TaskStartEvent): void {
  const task = event.execution.task;
  if (isOfficeTask(task) || isDebugPreLaunchTask(task)) {
    trackedTasks.add(task.name);
  }
}

function onDidEndTaskHandler(event: vscode.TaskEndEvent): void {
  if (isOfficeTask(event.execution.task) || isDebugPreLaunchTask(event.execution.task)) {
    trackedTasks.delete(event.execution.task.name);
  }
}

function onDidStartTaskProcessHandler(event: vscode.TaskProcessStartEvent): void {
  if (
    globalVariables.workspaceUri &&
    isValidOfficeAddInProject(globalVariables.workspaceUri.fsPath)
  ) {
    const task = event.execution.task;
    if (task.scope !== undefined && (isOfficeTask(task) || isDebugPreLaunchTask(task))) {
      allRunningOfficeTasks.set(getTaskKey(task), event.processId);
    }
  }
}

function onDidEndTaskProcessHandler(event: vscode.TaskProcessEndEvent): void {
  const task = event.execution.task;
  if (task.scope !== undefined && (isOfficeTask(task) || isDebugPreLaunchTask(task))) {
    allRunningOfficeTasks.delete(getTaskKey(task));
  }
}

async function onDidStartDebugSessionHandler(event: vscode.DebugSession): Promise<void> {
  if (
    globalVariables.workspaceUri &&
    isValidOfficeAddInProject(globalVariables.workspaceUri.fsPath)
  ) {
    const debugConfig = event.configuration;
    if (debugConfig && debugConfig.name && !debugConfig.postDebugTask) {
      if (await checkAndSkipDebugging()) {
        throw new Error(DebugSessionExists);
      } else {
        startLocalDebugSession();
      }
      allRunningDebugSessions.add(event.id);
    }

    await updateProjectStatus(
      globalVariables.workspaceUri.fsPath,
      CommandKey.LocalDebug,
      ok(undefined),
      true
    );
  }
}

function onDidTerminateDebugSessionHandler(event: vscode.DebugSession): void {
  if (allRunningDebugSessions.has(event.id)) {
    terminateAllRunningOfficeTasks();

    allRunningDebugSessions.delete(event.id);
    if (allRunningDebugSessions.size == 0) {
      endLocalDebugSession();
    }
    allRunningOfficeTasks.clear();
  }
}

export function terminateAllRunningOfficeTasks(): void {
  for (const task of allRunningOfficeTasks) {
    try {
      if (task[1] > 0) {
        process.kill(task[1], "SIGTERM");
      }
    } catch (e) {
      // ignore and keep killing others
    }
  }
  allRunningOfficeTasks.clear();
}

export function registerOfficeTaskAndDebugEvents(): void {
  globalVariables.context.subscriptions.push({
    dispose() {
      trackedTasks.clear();
    },
  });

  globalVariables.context.subscriptions.push(vscode.tasks.onDidStartTask(onDidStartTaskHandler));
  globalVariables.context.subscriptions.push(vscode.tasks.onDidEndTask(onDidEndTaskHandler));

  globalVariables.context.subscriptions.push(
    vscode.tasks.onDidStartTaskProcess((event: vscode.TaskProcessStartEvent) =>
      Correlator.runWithId(getLocalDebugSessionId(), onDidStartTaskProcessHandler, event)
    )
  );

  globalVariables.context.subscriptions.push(
    vscode.tasks.onDidEndTaskProcess((event: vscode.TaskProcessEndEvent) =>
      Correlator.runWithId(getLocalDebugSessionId(), onDidEndTaskProcessHandler, event)
    )
  );

  globalVariables.context.subscriptions.push(
    vscode.debug.onDidStartDebugSession((event: vscode.DebugSession) =>
      Correlator.runWithId(getLocalDebugSessionId(), onDidStartDebugSessionHandler, event)
    )
  );
  globalVariables.context.subscriptions.push(
    vscode.debug.onDidTerminateDebugSession((event: vscode.DebugSession) =>
      Correlator.runWithId(getLocalDebugSessionId(), onDidTerminateDebugSessionHandler, event)
    )
  );
}
