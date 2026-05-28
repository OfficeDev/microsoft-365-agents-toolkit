// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { performance } from "perf_hooks";

import { FxError, err, ok } from "@microsoft/teamsfx-api";
import {
  LocalEnvManager,
  LocalTelemetryReporter,
  TaskCommand,
  TaskLabel,
  TaskOverallLabel,
  TeamsFxNpmCommands,
} from "@microsoft/teamsfx-core";

import * as globalVariables from "../globalVariables";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryMeasurements,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { getLocalDebugSession } from "./common/localDebugSession";
import { updateProjectStatus } from "../utils/projectStatusUtils";
import { CommandKey } from "../constants";
import { TeamsFxTaskType } from "./common/debugConstants";
import detectPort from "detect-port";

function saveEventTime(eventName: string, time: number) {
  const session = getLocalDebugSession();
  // Assuming the event is only sent once in one local debug session,
  // because we only use the "high-level" events like debug-prerequisites, debug-precheck, etc..
  // And these events are indeed sent once.
  session.eventTimes[eventName] = time;
}

export const localTelemetryReporter = new LocalTelemetryReporter(
  {
    // Cannot directly refer to a global function because of import dependency cycle in ../telemetry/extTelemetry.ts.
    sendTelemetryEvent: (
      eventName: string,
      properties?: { [p: string]: string },
      measurements?: { [p: string]: number }
    ) => ExtTelemetry.sendTelemetryEvent(eventName, properties, measurements),

    sendTelemetryErrorEvent: (
      eventName: string,
      error: FxError,
      properties?: { [p: string]: string },
      measurements?: { [p: string]: number },
      errorProps?: string[]
    ) =>
      ExtTelemetry.sendTelemetryErrorEvent(eventName, error, properties, measurements, errorProps),
  },
  saveEventTime
);

export async function sendDebugAllStartEvent(additionalProperties: {
  [key: string]: string;
}): Promise<void> {
  const session = getLocalDebugSession();
  Object.assign(session.properties, additionalProperties);

  const properties = Object.assign(
    { [TelemetryProperty.CorrelationId]: session.id },
    session.properties
  );

  // Transparent task properties
  const taskInfo = await getTaskInfo();
  if (taskInfo && taskInfo.IsTransparentTask) {
    properties[TelemetryProperty.DebugPrelaunchTaskInfo] = JSON.stringify(
      taskInfo.PreLaunchTaskInfo
    );
    properties[TelemetryProperty.DebugIsTransparentTask] =
      properties[TelemetryProperty.DebugIsTransparentTask] ?? "true";
  } else {
    properties[TelemetryProperty.DebugIsTransparentTask] =
      properties[TelemetryProperty.DebugIsTransparentTask] ?? "false";
  }

  localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugAllStart, properties);
  return Promise.resolve();
}

export async function sendDebugAllEvent(
  error?: FxError,
  additionalProperties?: { [key: string]: string }
): Promise<void> {
  const session = getLocalDebugSession();
  const now = performance.now();

  if (globalVariables.workspaceUri?.fsPath) {
    await updateProjectStatus(
      globalVariables.workspaceUri.fsPath,
      CommandKey.LocalDebug,
      error ? err(error) : ok(undefined),
      true
    );
  }

  let duration = -1;
  const startTime = session.eventTimes[TelemetryEvent.DebugAllStart];
  if (startTime !== undefined) {
    duration = (now - startTime) / 1000;
  }

  // Calculate the time between two events
  // Event must be only once in one debug session.
  function durationBetween(eventStart: string, eventEnd: string): number {
    const t0 = session.eventTimes[eventStart];
    const t1 = session.eventTimes[eventEnd];
    if (t0 !== undefined && t1 !== undefined) {
      return t1 - t0;
    } else {
      return -1;
    }
  }

  // Calculate the "time gap" in a local debug session.
  // In current local debug implementation, there is some time that vscode is in control and extension has no control.
  // For example, between "debug-precheck" (task finishes) and "debug-all" (browser starts), vscode is starting the services.
  // However, we don't know when the services successfully start because we use problem matcher to determine the service start or fail.
  // And vscode does not provide a callback for that.
  // Estimating from the current data, this "time gap" can be up to 1 minute, so not neglectable.
  const precheckGap =
    durationBetween(TelemetryEvent.DebugPrerequisites, TelemetryEvent.DebugPreCheckStart) / 1000;
  const precheckTime = session.eventTimes[TelemetryEvent.DebugPreCheck];
  const servicesGap = precheckTime === undefined ? -1 : (performance.now() - precheckTime) / 1000;

  const properties: { [key: string]: string } = {
    [TelemetryProperty.CorrelationId]: session.id,
    [TelemetryProperty.Success]: error === undefined ? TelemetrySuccess.Yes : TelemetrySuccess.No,
    ...session.properties,
    ...additionalProperties,
  };

  // Transparent task properties
  const taskInfo = await getTaskInfo();
  if (taskInfo && taskInfo.IsTransparentTask) {
    properties[TelemetryProperty.DebugPrelaunchTaskInfo] = JSON.stringify(
      taskInfo.PreLaunchTaskInfo
    );
    properties[TelemetryProperty.DebugIsTransparentTask] =
      properties[TelemetryProperty.DebugIsTransparentTask] ?? "true";
  } else {
    properties[TelemetryProperty.DebugIsTransparentTask] =
      properties[TelemetryProperty.DebugIsTransparentTask] ?? "false";
  }

  const measurements = {
    [LocalTelemetryReporter.PropertyDuration]: duration,
    [TelemetryMeasurements.DebugPrecheckGapDuration]: precheckGap,
    [TelemetryMeasurements.DebugServicesGapDuration]: servicesGap,
  };

  const closedPorts: number[] = [];
  for (const port of globalVariables.LocalDebugPorts.checkPorts) {
    const port2 = await detectPort(port);
    if (port2 === port) {
      closedPorts.push(port);
    }
  }
  properties["debug-checked-ports"] = globalVariables.LocalDebugPorts.checkPorts.join(",");
  properties["debug-closed-ports"] = closedPorts.join(",");
  properties["debug-conflict-ports"] = globalVariables.LocalDebugPorts.conflictPorts.join(",");
  properties["debug-terminate-button"] = globalVariables.LocalDebugPorts.terminateButton;
  properties["debug-process2conflict-ports"] = JSON.stringify(
    globalVariables.LocalDebugPorts.process2conflictPorts
  );
  properties["debug-terminate-processes"] =
    globalVariables.LocalDebugPorts.terminateProcesses.join(",");

  if (error === undefined) {
    localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugAll, properties, measurements);
  } else {
    localTelemetryReporter.sendTelemetryErrorEvent(
      TelemetryEvent.DebugAll,
      error,
      properties,
      measurements
    );
  }
}

export const UnknownPlaceholder = "<unknown>";
export const UndefinedPlaceholder = "<undefined>";
export const DefaultPlaceholder = "<default>";

export function maskValue(
  value: string | undefined,
  knownValues: (string | { value: string; mask: string })[] = []
): string {
  if (typeof value === "undefined") {
    return UndefinedPlaceholder;
  }
  const findValue = knownValues.find((v) =>
    typeof v === "string" ? v === value : v.value === value
  );

  if (typeof findValue === "undefined") {
    return UnknownPlaceholder;
  } else if (typeof findValue === "string") {
    return findValue;
  } else {
    return findValue.mask;
  }
}

export function maskArrayValue(
  valueArr: string[] | undefined,
  knownValues: (string | { value: string; mask: string })[] = []
): string[] | string {
  if (typeof valueArr === "undefined") {
    return UndefinedPlaceholder;
  }

  return valueArr.map((v) => maskValue(`${v}`, knownValues));
}

interface ITaskJson {
  tasks?: ITask[];
}

interface ITask {
  label?: string;
  type?: string;
  command?: string;
  dependsOn?: string | string[];
}

interface IDependsOn {
  label: string;
  type: string;
  command: string;
}

type PreLaunchTaskInfo = { [key: string]: IDependsOn[] | undefined };
type TaskInfo = {
  PreLaunchTaskInfo: PreLaunchTaskInfo;
  IsTransparentTask: boolean;
};

export async function getTaskInfo(): Promise<TaskInfo | undefined> {
  try {
    if (!globalVariables.isTeamsFxProject || !globalVariables.workspaceUri?.fsPath) {
      return undefined;
    }

    const localEnvManager = new LocalEnvManager();
    const taskJson = (await localEnvManager.getTaskJson(
      globalVariables.workspaceUri.fsPath
    )) as ITaskJson;
    if (!taskJson) {
      return undefined;
    }

    const getDependsOn = (overallTaskLabel: string) => {
      const dependsOnArr: IDependsOn[] = [];
      const overallTask = findTask(taskJson, overallTaskLabel);
      if (!overallTask || !overallTask.dependsOn) {
        return undefined;
      }
      const labelList: string[] = Array.isArray(overallTask.dependsOn)
        ? overallTask.dependsOn
        : typeof overallTask.dependsOn === "string"
        ? [overallTask.dependsOn]
        : [];

      for (const label of labelList) {
        const task = findTask(taskJson, label);
        const isTeamsFxTask =
          task?.type === TeamsFxTaskType ||
          (task?.type === "shell" &&
            task?.command &&
            Object.values(TeamsFxNpmCommands).includes(task?.command as any));

        // Only send the info scaffold by Teams Toolkit. If user changed some property, the value will be "unknown".
        dependsOnArr.push({
          label: maskValue(label, Object.values(TaskLabel)),
          type: maskValue(task?.type, [TeamsFxTaskType]),
          command: !isTeamsFxTask
            ? task?.command
              ? UnknownPlaceholder
              : UndefinedPlaceholder
            : maskValue(task?.command, [
                ...Object.values(TaskCommand),
                ...Object.values(TeamsFxNpmCommands),
              ]),
        });
      }
      return dependsOnArr;
    };
    const prelaunchTaskInfo: { [key: string]: IDependsOn[] | undefined } = {};
    Object.values(TaskOverallLabel).forEach((l) => {
      const dependsOn = getDependsOn(l);
      if (dependsOn) {
        prelaunchTaskInfo[l] = dependsOn;
      }
    });

    const teamsfxTasks = taskJson?.tasks?.filter(
      (t) =>
        t?.type === TeamsFxTaskType &&
        t?.command &&
        Object.values(TaskCommand).includes(t?.command as any)
    );

    return {
      PreLaunchTaskInfo: prelaunchTaskInfo,
      IsTransparentTask: !!teamsfxTasks?.length,
    };
  } catch {}

  return undefined;
}

function findTask(taskJson: ITaskJson, label: string): ITask | undefined {
  return taskJson?.tasks?.find((task) => task?.label === label);
}
