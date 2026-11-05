// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Err, FxError, M365TokenProvider, Platform, Tools, err } from "@microsoft/teamsfx-api";
import path from "path";
import type { DriverContext } from "../component/driver/interface/commonArgs";
import type { ExecutionResult, StepDriver } from "../component/driver/interface/stepDriver";
import { errorHandle } from "../component/utils/common";
import {
  ActionArgs,
  ActionName,
  ActionNames,
  createActionError,
  getActionDefinition,
  validateActionArgs,
} from "./contracts";

export interface ActionExecutionOptions {
  projectPath: string;
  signal?: AbortSignal;
}

const drivers = {
  [ActionNames.Acquire]: async () => {
    const { M365TitleAcquireDriver } = await import("../component/driver/m365/acquire");
    return new M365TitleAcquireDriver();
  },
  [ActionNames.PublishPolicy]: async () => {
    const { PublishPolicyDriver } = await import("../component/driver/workiq/publishPolicy");
    return new PublishPolicyDriver();
  },
} satisfies Record<ActionName, () => Promise<StepDriver>>;

export type DriverExecutionResult<TDriver extends StepDriver> = Awaited<
  ReturnType<TDriver["execute"]>
>;

type RegisteredActionResults = {
  [K in ActionName]: DriverExecutionResult<Awaited<ReturnType<(typeof drivers)[K]>>>;
};

export interface ActionExecutionFailure {
  result: Err<Map<string, string>, FxError>;
  summaries: string[];
}

export type ActionResult<K extends ActionName> =
  RegisteredActionResults[K] | ActionExecutionFailure;

function silentProvider(provider: M365TokenProvider): M365TokenProvider {
  return {
    getAccessToken: (request) => provider.getAccessToken({ ...request, showDialog: false }),
    getJsonObject: (request, tenantId) =>
      provider.getJsonObject({ ...request, showDialog: false }, tenantId),
    getStatus: (request) => provider.getStatus({ ...request, showDialog: false }),
    signout: () => provider.signout(),
    switchTenant: (tenantId) => provider.switchTenant(tenantId),
    setStatusChangeMap: (name, request, callback, immediateCall) =>
      provider.setStatusChangeMap(name, { ...request, showDialog: false }, callback, immediateCall),
    removeStatusChangeMap: (name) => provider.removeStatusChangeMap(name),
  };
}

export function executeAction<K extends ActionName>(
  tools: Tools,
  name: K,
  args: ActionArgs<NoInfer<K>>,
  options: ActionExecutionOptions
): Promise<ActionResult<K>>;
export async function executeAction(
  tools: Tools,
  name: ActionName,
  args: unknown,
  options: ActionExecutionOptions
): Promise<ExecutionResult> {
  const definition = getActionDefinition(name);
  if (definition.isErr()) return { result: err(definition.error), summaries: [] };
  const checked = validateActionArgs(name, args);
  if (checked.isErr()) return { result: err(checked.error), summaries: [] };
  const projectPath = options?.projectPath;
  const signal = options?.signal;
  if (
    typeof projectPath !== "string" ||
    !path.isAbsolute(projectPath) ||
    (signal !== undefined && !(signal instanceof AbortSignal))
  ) {
    return { result: err(createActionError("InvalidActionContext")), summaries: [] };
  }
  if (signal && !definition.value.supportsCancellation) {
    return { result: err(createActionError("UnsupportedCancellation")), summaries: [] };
  }
  const context: DriverContext = {
    projectPath,
    platform: Platform.CLI,
    nonInteractive: true,
    signal,
    m365TokenProvider: silentProvider(tools.tokenProvider.m365TokenProvider),
    azureAccountProvider: tools.tokenProvider.azureAccountProvider,
    logProvider: tools.logProvider,
    telemetryReporter: tools.telemetryReporter,
    ui: undefined,
    progressBar: undefined,
  };
  try {
    const driver: StepDriver = await drivers[name]();
    return await driver.execute(checked.value, context);
  } catch (error) {
    return { result: await errorHandle(error, "executeAction"), summaries: [] };
  }
}
