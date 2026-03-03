// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, IProgressHandler, ok, SystemError, UserError } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { BaseComponentInnerError } from "../../error/componentError";
import { TeamsFxTelemetryReporter } from "../../utils/teamsFxTelemetryReporter";
import { logMessageKeys } from "../aad/utility/constants";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult } from "../interface/stepDriver";

type ActionResult = ExecutionResult | ExecutionResult["result"];

// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export interface WrapDriverContext extends DriverContext {
  createProgressBar(title: string, steps: number): Promise<IProgressHandler | undefined>;
  addTelemetryProperties(properties: Record<string, string>): void;
  addSummary(...summaries: string[]): void;
}

// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export class WrapDriverContext {
  progressBars: IProgressHandler[] = [];
  summaries: string[] = [];
  eventName: string;
  telemetryProperties: Record<string, string>;
  wrapTelemetryReporter?: TeamsFxTelemetryReporter;
  constructor(driverContext: DriverContext, eventName: string, componentName: string) {
    Object.assign(this, driverContext, {});
    this.eventName = eventName;
    this.telemetryProperties = {
      component: componentName.replace(/\//g, ""), // Remove `/` in the componentName to avoid the value being redacted.
    };
    if (driverContext.telemetryReporter) {
      this.wrapTelemetryReporter = new TeamsFxTelemetryReporter(driverContext.telemetryReporter, {
        componentName,
      });
    }
  }

  async createProgressBar(title: string, steps: number): Promise<IProgressHandler | undefined> {
    const progressBar = this.ui?.createProgressBar(title, steps);
    if (progressBar) {
      this.progressBars.push(progressBar);
      await progressBar.start();
      await progressBar.next();
    }
    return progressBar;
  }

  async endProgressBars(success: boolean): Promise<void> {
    await Promise.all(
      this.progressBars.map(async (progressbar) => {
        await progressbar.end(success);
      })
    );
  }

  addTelemetryProperties(properties: Record<string, string>): void {
    this.telemetryProperties = { ...properties, ...this.telemetryProperties };
  }

  addSummary(...summaries: string[]): void {
    this.summaries.push(...summaries);
  }
}

export async function wrapRun(
  context: WrapDriverContext,
  exec: () => Promise<Map<string, string>>,
  isExecute?: boolean
): Promise<ActionResult> {
  const eventName = context.eventName;
  try {
    let actionRes: ActionResult;
    context.wrapTelemetryReporter?.sendStartEvent({ eventName });
    context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, eventName));
    const res = await exec();
    context.wrapTelemetryReporter?.sendEndEvent({
      eventName,
      properties: context.telemetryProperties,
    });
    context.logProvider?.info(getLocalizedString(logMessageKeys.successExecuteDriver, eventName));
    await context.endProgressBars(true);
    if (isExecute) {
      actionRes = { result: ok(res), summaries: context.summaries };
    } else {
      actionRes = ok(res);
    }
    return actionRes;
  } catch (error: any) {
    let actionRes: ActionResult;
    const fxError = getError(context, error);
    context.wrapTelemetryReporter?.sendEndEvent(
      {
        eventName,
        properties: context.telemetryProperties,
      },
      fxError
    );
    await context.endProgressBars(false);
    if (isExecute) {
      actionRes = { result: err(fxError), summaries: context.summaries };
    } else {
      actionRes = err(fxError);
    }
    return actionRes;
  }
}
const ErrorConstants = {
  unhandledError: "UnhandledError",
  unhandledErrorMessage: "Unhandled Error",
};
function getError(context: WrapDriverContext, error: any): FxError {
  let fxError: FxError;
  if (error instanceof BaseComponentInnerError) {
    fxError = error.toFxError();
  } else if (error instanceof UserError || error instanceof SystemError) {
    fxError = error;
  } else {
    if (!(error instanceof Error)) {
      error = new Error(error.toString());
    }
    fxError = new SystemError({
      error,
      source: context.eventName,
      name: ErrorConstants.unhandledError,
      message: error.message,
      displayMessage: error.message,
    });
  }
  return fxError;
}
