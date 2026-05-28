// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, Void, FxError, err, ok, UserError, SystemError } from "@microsoft/teamsfx-api";
import { DepsType, Prerequisite, TelemetryContext } from "@microsoft/teamsfx-core";
import { getLocalDebugSession } from "../common/localDebugSession";
import { v3PrerequisiteCheckTaskDisplayMessages } from "../common/debugConstants";
import { localTelemetryReporter } from "../localTelemetryReporter";
import { terminateAllRunningTeamsfxTasks } from "../teamsfxTaskHandler";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { Checker } from "./prerequisitesCheckerConstants";
import { CheckResult, PrerequisiteOrderedChecker } from "../common/types";
import VsCodeLogInstance from "../../commonlib/log";
import { _checkAndInstall } from "./common";
import { allRunningTeamsfxTasks } from "../common/globalVariables";

export async function checkAndInstallForTask(
  prerequisites: string[],
  ports: number[] | undefined,
  telemetryProperties: { [key: string]: string }
): Promise<Result<Void, FxError>> {
  const orderedCheckers = getOrderedCheckersForTask(prerequisites, ports);

  const additionalTelemetryProperties = Object.assign(
    {
      [TelemetryProperty.DebugIsTransparentTask]: "true",
    },
    telemetryProperties
  );
  return await localTelemetryReporter.runWithTelemetryProperties(
    TelemetryEvent.DebugPrerequisites,
    additionalTelemetryProperties,
    async (ctx: TelemetryContext) => {
      // terminate all running teamsfx tasks
      if (allRunningTeamsfxTasks.size > 0) {
        VsCodeLogInstance.info("Terminate all running teamsfx tasks.");
        await terminateAllRunningTeamsfxTasks();
      }

      const res = await _checkAndInstall(
        v3PrerequisiteCheckTaskDisplayMessages,
        orderedCheckers,
        additionalTelemetryProperties
      );
      if (res.error) {
        const debugSession = getLocalDebugSession();
        addCheckResultsForTelemetry(
          res.checkResults,
          debugSession.properties,
          debugSession.errorProps
        );
        addCheckResultsForTelemetry(res.checkResults, ctx.properties, ctx.errorProps);
        return err(res.error);
      }
      return ok(Void);
    }
  );
}

function getOrderedCheckersForTask(
  prerequisites: string[],
  ports?: number[]
): PrerequisiteOrderedChecker[] {
  const checkers: PrerequisiteOrderedChecker[] = [];
  if (prerequisites.includes(Prerequisite.nodejs)) {
    checkers.push({ info: { checker: DepsType.ProjectNode }, fastFail: true });
  }
  if (prerequisites.includes(Prerequisite.m365Account)) {
    checkers.push({ info: { checker: Checker.M365Account }, fastFail: false });
  }
  if (prerequisites.includes(Prerequisite.copilotAccess)) {
    checkers.push({ info: { checker: Checker.CopilotAccess }, fastFail: false });
  }
  if (prerequisites.includes(Prerequisite.portOccupancy)) {
    checkers.push({ info: { checker: Checker.Ports, ports: ports }, fastFail: false });
  }
  return checkers;
}

function addCheckResultsForTelemetry(
  checkResults: CheckResult[],
  properties: { [key: string]: string },
  errorProps: string[]
): void {
  // const [resultRaw, resultSafe] = convertCheckResultsForTelemetry(checkResults);
  // properties[TelemetryProperty.DebugCheckResultsSafe] = resultSafe;
  // properties[TelemetryProperty.DebugCheckResults] = maskSecret(resultRaw, { replace: "***" });
  // only the raw event contains error message
  errorProps.push(TelemetryProperty.DebugCheckResults);
}

// Mainly addresses two issues:
// 1. Some error messages contain special characters which will cause the whole debug-check-results to be redacted.
// 2. CheckResult[] is hard to parse in kusto query (an array of objects).
//
// `debug-check-results` contains only known content and we know it will not be redacted.
// `debug-check-results-raw` might contain arbitrary string and be redacted.
function convertCheckResultsForTelemetry(checkResults: CheckResult[]): [string, string] {
  const resultRaw: { [checker: string]: unknown } = {};
  const resultSafe: { [checker: string]: { [key: string]: string | undefined } } = {};
  for (const checkResult of checkResults) {
    resultRaw[checkResult.checker] = checkResult;
    resultSafe[checkResult.checker] = {
      result: checkResult.result,
      source: checkResult.error?.source,
      errorCode: checkResult.error?.name,
      errorType:
        checkResult.error === undefined
          ? undefined
          : checkResult.error instanceof UserError
          ? "user"
          : checkResult.error instanceof SystemError
          ? "system"
          : "unknown",
    };
  }

  return [JSON.stringify(resultRaw), JSON.stringify(resultSafe)];
}
