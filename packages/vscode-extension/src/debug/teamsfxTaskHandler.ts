// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ProductName, UserError } from "@microsoft/teamsfx-api";
import {
  Correlator,
  getHashedEnv,
  Hub,
  isValidProject,
  TaskCommand,
} from "@microsoft/teamsfx-core";
import path from "path";
import { performance } from "perf_hooks";
import * as util from "util";
import * as vscode from "vscode";
import VsCodeLogInstance from "../commonlib/log";
import { ExtensionErrors, ExtensionSource } from "../error/error";
import * as globalVariables from "../globalVariables";
import { cdpClientManager, isM365CopilotChatDebugConfiguration } from "../pluginDebugger/cdpClient";
import { VS_CODE_UI } from "../qm/vsc_ui";
import {
  TelemetryEvent,
  TelemetryMeasurements,
  TelemetryProperty,
} from "../telemetry/extTelemetryEvents";
import { getNpmInstallLogInfo, getTestToolLogInfo } from "../utils/localEnvManagerUtils";
import { localize } from "../utils/localizeUtils";
import { processUtil } from "../utils/processUtil";
import {
  clearAADAfterLocalDebugHelpLink,
  DebugNoSessionId,
  errorDetail,
  issueChooseLink,
  issueLink,
  issueTemplate,
  m365AppsPrerequisitesHelpLink,
} from "./common/debugConstants";
import { allRunningTeamsfxTasks } from "./common/globalVariables";
import {
  endLocalDebugSession,
  getLocalDebugSession,
  getLocalDebugSessionId,
} from "./common/localDebugSession";
import { TeamsfxDebugConfiguration } from "./common/teamsfxDebugConfiguration";
import { deleteAad } from "./deleteAadHelper";
import { localTelemetryReporter, sendDebugAllEvent } from "./localTelemetryReporter";
import { allRunningDebugSessions } from "./officeTaskHandler";
import { BaseTunnelTaskTerminal } from "./taskTerminal/baseTunnelTaskTerminal";

class NpmInstallTaskInfo {
  private startTime: number;

  constructor() {
    this.startTime = performance.now();
  }

  public getDurationSecondsUntilNow() {
    return (performance.now() - this.startTime) / 1000;
  }
}

const activeNpmInstallTasks = new Map<string, NpmInstallTaskInfo>();

/**
 * This EventEmitter is used to track all running tasks called by `runTask`.
 * Each task executed by `runTask` will have an internal task id.
 * Event emitters use this id to identify each tracked task, and `runTask` matches this id
 * to determine whether a task is terminated or not.
 */
let taskEndEventEmitter: vscode.EventEmitter<{
  id: string;
  name: string;
  exitCode?: number;
}>;
let taskStartEventEmitter: vscode.EventEmitter<string>;
const trackedTasks = new Map<string, string>();

function getTaskKey(task: vscode.Task): string {
  if (task === undefined) {
    return "";
  }

  // "source|name|scope"
  const scope = (task.scope as vscode.WorkspaceFolder)?.uri?.toString() || task.scope?.toString();
  return `${task.source}|${task.name}|${scope ?? ""}`;
}

function isNpmInstallTask(task: vscode.Task): boolean {
  if (task) {
    return task.name.trim().toLocaleLowerCase().endsWith("npm install");
  }

  return false;
}

function isCheckDevProxyTask(task: vscode.Task): boolean {
  if (task.definition.type === "shell" && task.execution && <vscode.ShellExecution>task.execution) {
    const execution = <vscode.ShellExecution>task.execution;
    return (
      execution.options?.cwd === "${workspaceFolder}/proxy" &&
      execution.commandLine === "node check.js"
    );
  }
  return false;
}

function isTeamsFxTransparentTask(task: vscode.Task): boolean {
  if (task.definition && task.definition.type === ProductName) {
    const command = task.definition.command as string;
    if (Object.values(TaskCommand).includes(command as any)) {
      return true;
    }
  }
  return false;
}

function isTeamsfxTask(task: vscode.Task): boolean {
  // teamsfx: xxx start / xxx watch
  if (task) {
    if (
      task.source === ProductName &&
      (task.name.trim().toLocaleLowerCase().endsWith("start") ||
        task.name.trim().toLocaleLowerCase().endsWith("watch"))
    ) {
      // provided by toolkit
      return true;
    }

    if (task.definition && task.definition.type === ProductName) {
      // defined by launch.json
      const command = task.definition.command as string;
      return (
        command !== undefined &&
        (command.trim().toLocaleLowerCase().endsWith("start") ||
          command.trim().toLocaleLowerCase().endsWith("watch"))
      );
    }

    // dev:teamsfx and watch:teamsfx
    let commandLine: string | undefined;
    if (task.execution && <vscode.ShellExecution>task.execution) {
      const execution = <vscode.ShellExecution>task.execution;
      commandLine =
        execution.commandLine ||
        `${typeof execution.command === "string" ? execution.command : execution.command.value} ${(
          execution.args || []
        ).join(" ")}`;
    }
    if (commandLine !== undefined) {
      if (/(npm|yarn)[\s]+(run )?[\s]*[^:\s]+:teamsfx/i.test(commandLine)) {
        return true;
      }
      if (/teamsfx\/script\/.*\.js/i.test(commandLine)) {
        return true;
      }
    }
    if (isCheckDevProxyTask(task)) {
      return true;
    }
  }

  return false;
}

function isLaunchTestToolTask(task: vscode.Task): boolean {
  if (task) {
    // dev:teamsfx and watch:teamsfx
    if (task.execution && <vscode.ShellExecution>task.execution) {
      const execution = <vscode.ShellExecution>task.execution;
      const commandLine =
        execution.commandLine ||
        `${typeof execution.command === "string" ? execution.command : execution.command.value} ${(
          execution.args || []
        ).join(" ")}`;
      if (/(npm|yarn)[\s]+(run )?[\s]*[^:\s]+:teamsfx:launch-testtool/i.test(commandLine)) {
        return true;
      }
    }
  }

  return false;
}

function displayTerminal(taskName: string): boolean {
  const terminal = vscode.window.terminals.find((t) => t.name === taskName);
  if (terminal !== undefined && terminal !== vscode.window.activeTerminal) {
    terminal.show(true);
    return true;
  }

  return false;
}

function onDidStartTaskHandler(event: vscode.TaskStartEvent): void {
  const taskId = event.execution.task?.definition?.teamsfxTaskId;
  if (taskId !== undefined) {
    trackedTasks.set(taskId, event.execution.task.name);
    taskStartEventEmitter.fire(taskId as string);
  }
}

function onDidEndTaskHandler(event: vscode.TaskEndEvent): void {
  const taskId = event.execution.task?.definition?.teamsfxTaskId;
  if (taskId !== undefined && trackedTasks.has(taskId as string)) {
    trackedTasks.delete(taskId as string);
    taskEndEventEmitter.fire({
      id: taskId as string,
      name: event.execution.task.name,
      exitCode: undefined,
    });
  }
}

async function onDidStartTaskProcessHandler(event: vscode.TaskProcessStartEvent): Promise<void> {
  if (globalVariables.workspaceUri && isValidProject(globalVariables.workspaceUri.fsPath)) {
    const task = event.execution.task;
    if (task.scope !== undefined && isTeamsfxTask(task)) {
      allRunningTeamsfxTasks.set(getTaskKey(task), event.processId);
      localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugServiceStart, {
        [TelemetryProperty.DebugServiceName]: task.name,
      });
    } else if (isNpmInstallTask(task)) {
      localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugNpmInstallStart, {
        [TelemetryProperty.DebugNpmInstallName]: task.name,
      });

      activeNpmInstallTasks.set(task.name, new NpmInstallTaskInfo());
    }
    if (isLaunchTestToolTask(task)) {
      // Debug in Test Tool doesn't launch browser in vscode launch.json
      // This is the final task that we have control to send debug-all event.
      const session = getLocalDebugSession();
      if (session.id !== DebugNoSessionId) {
        if (session.failedServices.length > 0) {
          await terminateAllRunningTeamsfxTasks();
          // TODO: send test tool log file
          await sendDebugAllEvent(
            new UserError({
              source: ExtensionSource,
              name: ExtensionErrors.DebugServiceFailedBeforeStartError,
            }),
            {
              [TelemetryProperty.DebugFailedServices]: JSON.stringify(session.failedServices),
              [TelemetryProperty.DebugTestTool]: "true",
            }
          );
          endLocalDebugSession();
          return;
        }
        await sendDebugAllEvent(undefined, { [TelemetryProperty.DebugTestTool]: "true" });
        // Need to endLocalDebugSession() here.
        // Otherwise, when user stops debugging, or task exits, it will incorrectly send an incorrect debug-all event with success=no
        endLocalDebugSession();
      }
    }
  }
}

async function onDidEndTaskProcessHandler(event: vscode.TaskProcessEndEvent): Promise<void> {
  const timestamp = new Date();
  const task = event.execution.task;
  const activeTerminal = vscode.window.activeTerminal;

  const taskId = task?.definition?.teamsfxTaskId;
  if (taskId !== undefined) {
    trackedTasks.delete(taskId as string);
    taskEndEventEmitter.fire({
      id: taskId as string,
      name: event.execution.task.name,
      exitCode: event.exitCode,
    });
  }

  if (task.scope !== undefined && isTeamsfxTask(task)) {
    const currentSession = getLocalDebugSession();
    if (event.exitCode !== 0) {
      currentSession.failedServices.push({ name: task.name, exitCode: event.exitCode });
    }
    allRunningTeamsfxTasks.delete(getTaskKey(task));
    localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugService, {
      [TelemetryProperty.DebugServiceName]: task.name,
      [TelemetryProperty.DebugServiceExitCode]: String(event.exitCode),
    });

    if (isLaunchTestToolTask(task)) {
      // This is the final task for test tool.
      // If this task exits (even exitCode is 0) before being successfully started, the debug fails.
      if (currentSession.id !== DebugNoSessionId) {
        await terminateAllRunningTeamsfxTasks();
        await sendDebugAllEvent(
          new UserError({
            source: ExtensionSource,
            name: ExtensionErrors.DebugTestToolFailedToStartError,
          }),
          {
            [TelemetryProperty.DebugTestTool]: "true",
            [TelemetryProperty.DebugServiceExitCode]: String(event.exitCode),
            [TelemetryProperty.DebugTestToolLog]: (await getTestToolLogInfo()) || "<undefined>",
          }
        );
        endLocalDebugSession();
      }
    } else {
      vscode.window
        .showInformationMessage(
          localize("teamstoolkit.localDebug.deleteAADNotification"),
          localize("teamstoolkit.localDebug.learnMore")
        )
        .then(
          async (result) => {
            if (result === localize("teamstoolkit.localDebug.learnMore")) {
              await VS_CODE_UI.openUrl(clearAADAfterLocalDebugHelpLink);
            }
          },
          () => {
            // Do nothing on reject
          }
        );
    }
  } else if (
    task.scope !== undefined &&
    isTeamsFxTransparentTask(task) &&
    event.exitCode !== 0 &&
    event.exitCode !== -1
  ) {
    await terminateAllRunningTeamsfxTasks();
  } else if (isNpmInstallTask(task)) {
    try {
      const taskInfo = activeNpmInstallTasks.get(task.name);
      activeNpmInstallTasks.delete(task.name);
      if (activeTerminal?.name === task.name && event.exitCode === 0) {
        // when the task in active terminal is ended successfully.
        for (const hiddenTaskName of activeNpmInstallTasks.keys()) {
          // display the first hidden terminal.
          if (displayTerminal(hiddenTaskName)) {
            return;
          }
        }
      } else if (activeTerminal?.name !== task.name && event.exitCode !== 0) {
        // when the task in hidden terminal failed to execute.
        displayTerminal(task.name);
      }

      const cwdOption = (task.execution as vscode.ShellExecution).options?.cwd;
      let cwd: string | undefined;
      if (cwdOption !== undefined) {
        cwd = cwdOption.replace("${workspaceFolder}", globalVariables.workspaceUri!.fsPath);
      }
      const npmInstallLogInfo = await getNpmInstallLogInfo();
      let validNpmInstallLogInfo = false;
      if (
        cwd !== undefined &&
        npmInstallLogInfo?.cwd !== undefined &&
        path.relative(npmInstallLogInfo.cwd, cwd).length === 0 &&
        event.exitCode !== undefined &&
        npmInstallLogInfo.exitCode === event.exitCode
      ) {
        const timeDiff = timestamp.getTime() - npmInstallLogInfo.timestamp.getTime();
        if (timeDiff >= 0 && timeDiff <= 20000) {
          validNpmInstallLogInfo = true;
        }
      }
      const properties: { [key: string]: string } = {
        [TelemetryProperty.DebugNpmInstallName]: task.name,
        [TelemetryProperty.DebugNpmInstallExitCode]: String(event.exitCode), // "undefined" or number value
      };
      if (validNpmInstallLogInfo) {
        properties[TelemetryProperty.DebugNpmInstallNodeVersion] = String(
          npmInstallLogInfo?.nodeVersion
        ); // "undefined" or string value
        properties[TelemetryProperty.DebugNpmInstallNpmVersion] = String(
          npmInstallLogInfo?.npmVersion
        ); // "undefined" or string value
        properties[TelemetryProperty.DebugNpmInstallErrorMessage] = String(
          npmInstallLogInfo?.errorMessage?.join("\n")
        ); // "undefined" or string value
      }

      const measurements: { [key: string]: number } = {};
      if (taskInfo !== undefined) {
        measurements[TelemetryMeasurements.Duration] = taskInfo.getDurationSecondsUntilNow();
      }

      if (event.exitCode !== 0 || properties[TelemetryProperty.DebugNpmInstallErrorMessage]) {
        localTelemetryReporter.sendTelemetryErrorEvent(
          TelemetryEvent.DebugNpmInstall,
          new UserError({ name: ExtensionErrors.DebugNpmInstallError, source: ExtensionSource }),
          properties,
          measurements,
          [TelemetryProperty.DebugNpmInstallErrorMessage]
        );
      } else {
        localTelemetryReporter.sendTelemetryEvent(
          TelemetryEvent.DebugNpmInstall,
          properties,
          measurements
        );
      }

      if (cwd !== undefined && event.exitCode !== undefined && event.exitCode !== 0) {
        // Do not show this hint message for prerequisites check and automatic npm install
        if (taskId === undefined) {
          let url: string;
          if (validNpmInstallLogInfo) {
            url = `${issueLink}title=new+bug+report: Task '${
              task.name
            }' failed&body=${issueTemplate}${errorDetail}${JSON.stringify(
              npmInstallLogInfo,
              undefined,
              4
            )}`;
          } else {
            url = issueChooseLink;
          }
          const issue = {
            title: localize("teamstoolkit.handlers.reportIssue"),
            run: async (): Promise<void> => {
              await vscode.commands.executeCommand("vscode.open", vscode.Uri.parse(url));
            },
          };
          vscode.window
            .showErrorMessage(
              util.format(
                localize("teamstoolkit.localDebug.npmInstallFailedHintMessage"),
                task.name,
                task.name
              ),
              issue
            )
            .then(
              async (button) => {
                await button?.run();
              },
              () => {
                // Do nothing on reject
              }
            );
          VsCodeLogInstance.error(
            util.format(
              localize("teamstoolkit.localDebug.npmInstallFailedHintMessage"),
              task.name,
              task.name
            )
          );
        }
        await terminateAllRunningTeamsfxTasks();
      }
    } catch {
      // ignore any error
    }
  }
}

async function onDidStartDebugSessionHandler(event: vscode.DebugSession): Promise<void> {
  const port = isM365CopilotChatDebugConfiguration(event.configuration);
  if (port) {
    const url = event.configuration.url;
    const name = event.configuration.name;
    cdpClientManager.start(url, port, name);
  }
  if (globalVariables.workspaceUri && isValidProject(globalVariables.workspaceUri.fsPath)) {
    const debugConfig = event.configuration as TeamsfxDebugConfiguration;
    if (
      debugConfig &&
      debugConfig.name &&
      (debugConfig.url || debugConfig.port) && // it's from launch.json
      !debugConfig.postRestartTask
    ) {
      allRunningDebugSessions.add(event.id);

      // show M365 tenant hint message for Outlook and Office
      if (debugConfig.teamsfxHub === Hub.outlook || debugConfig.teamsfxHub === Hub.office) {
        VS_CODE_UI.showMessage(
          "info",
          localize("teamstoolkit.localDebug.m365TenantHintMessage"),
          false,
          localize("teamstoolkit.localDebug.learnMore")
        ).then(
          async (result) => {
            if (result.isOk() && result.value === localize("teamstoolkit.localDebug.learnMore")) {
              await VS_CODE_UI.openUrl(m365AppsPrerequisitesHelpLink);
            }
          },
          () => {
            // Do nothing on reject
          }
        );
      }

      // and not a restart one
      // send f5 event telemetry
      let env = "";
      if (debugConfig.teamsfxEnv) {
        env = getHashedEnv(debugConfig.teamsfxEnv);
      }

      localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugStart, {
        [TelemetryProperty.DebugSessionId]: event.id,
        [TelemetryProperty.DebugType]: debugConfig.type,
        [TelemetryProperty.DebugRequest]: debugConfig.request,
        [TelemetryProperty.DebugPort]: String(debugConfig.port),
        [TelemetryProperty.DebugRemote]: String(debugConfig.teamsfxIsRemote),
        [TelemetryProperty.DebugAppId]: String(debugConfig.teamsfxAppId),
        [TelemetryProperty.Env]: env,
        [TelemetryProperty.Hub]: String(debugConfig.teamsfxHub),
      });
      // This is the launch browser local debug session.
      if (debugConfig.request === "launch" && !debugConfig.teamsfxIsRemote) {
        // Handle cases that some services failed immediately after start.
        const currentSession = getLocalDebugSession();
        if (currentSession.id !== DebugNoSessionId && currentSession.failedServices.length > 0) {
          await terminateAllRunningTeamsfxTasks();
          await vscode.debug.stopDebugging();
          await sendDebugAllEvent(
            new UserError({
              source: ExtensionSource,
              name: ExtensionErrors.DebugServiceFailedBeforeStartError,
            }),
            {
              [TelemetryProperty.DebugFailedServices]: JSON.stringify(
                currentSession.failedServices
              ),
            }
          );
          endLocalDebugSession();
          return;
        }

        sendDebugAllEvent().catch(() => {
          // Do nothing
        });
      }
    }
  }
}

export async function terminateAllRunningTeamsfxTasks(): Promise<void> {
  for (const task of allRunningTeamsfxTasks) {
    try {
      if (task[1] > 0) {
        await processUtil.killProcess(task[1]);
      }
    } catch (e) {
      // ignore and keep killing others
    }
  }
  allRunningTeamsfxTasks.clear();
  void deleteAad();
  BaseTunnelTaskTerminal.stopAll();
}

async function onDidTerminateDebugSessionHandler(event: vscode.DebugSession): Promise<void> {
  const port = isM365CopilotChatDebugConfiguration(event.configuration);
  if (port) {
    await cdpClientManager.stop(port);
  }
  if (allRunningDebugSessions.has(event.id)) {
    // a valid debug session
    // send stop-debug event telemetry
    localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugStop, {
      [TelemetryProperty.DebugSessionId]: event.id,
    });

    await terminateAllRunningTeamsfxTasks();

    allRunningDebugSessions.delete(event.id);
    if (allRunningDebugSessions.size == 0) {
      endLocalDebugSession();
    }
    allRunningTeamsfxTasks.clear();
  }
}

export function registerTeamsfxTaskAndDebugEvents(): void {
  taskEndEventEmitter = new vscode.EventEmitter<{ id: string; name: string; exitCode?: number }>();
  taskStartEventEmitter = new vscode.EventEmitter<string>();
  globalVariables.context.subscriptions.push({
    dispose() {
      taskEndEventEmitter.dispose();
      taskStartEventEmitter.dispose();
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

  // debug session handler use correlation-id from event.configuration.teamsfxCorrelationId
  // to minimize concurrent debug session affecting correlation-id
  globalVariables.context.subscriptions.push(
    vscode.debug.onDidStartDebugSession((event: vscode.DebugSession) =>
      Correlator.runWithId(
        // fallback to retrieve correlation id from the global variable.
        event.configuration.teamsfxCorrelationId || getLocalDebugSessionId(),
        onDidStartDebugSessionHandler,
        event
      )
    )
  );
  globalVariables.context.subscriptions.push(
    vscode.debug.onDidTerminateDebugSession((event: vscode.DebugSession) =>
      Correlator.runWithId(
        event.configuration.teamsfxCorrelationId || getLocalDebugSessionId(),
        onDidTerminateDebugSessionHandler,
        event
      )
    )
  );
}
