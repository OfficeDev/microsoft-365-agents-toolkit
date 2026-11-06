// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ActionNames,
  type ActionArgs,
  type ActionResult,
  type AcquireArgs as PublicAcquireArgs,
  type PublishPolicyArgs as PublicPublishPolicyArgs,
} from "../../src/actions";
import type { AcquireArgs } from "../../src/component/driver/m365/args";
import type { FxCore } from "../../src/core/FxCore";
import type { IFxCoreClient } from "../../src/core/FxCoreClient";
import type { ExecutionResult, StepDriver } from "../../src/component/driver/interface/stepDriver";
import type { DriverContext } from "../../src/component/driver/interface/commonArgs";
import type { ActionExecutionFailure, DriverExecutionResult } from "../../src/actions/execute";
import type { PolicyPublicationReport } from "../../src/component/workiq/policyPublication";

export async function checkPolicyResultInference(
  client: IFxCoreClient,
  core: FxCore
): Promise<void> {
  const options = { projectPath: "C:\\project", signal: new AbortController().signal };
  const args: ActionArgs<typeof ActionNames.PublishPolicy> = {
    titleId: "U_title",
    policyPath: "policy.json",
    callerId: "11111111-2222-4333-8444-555555555555",
  };
  const publicArgs: PublicPublishPolicyArgs = args;
  const canonicalPolicyName: "copilotAgent/publishPolicy" = ActionNames.PublishPolicy;
  void client.executeAction(canonicalPolicyName, publicArgs, options);
  // @ts-expect-error The previous action spelling is no longer allowlisted.
  void client.executeAction("workiq/publishPolicy", publicArgs, options);
  const fromCore: ActionResult<typeof ActionNames.PublishPolicy> = await core.executeAction(
    ActionNames.PublishPolicy,
    publicArgs,
    options
  );
  const fromClient = await client.executeAction(ActionNames.PublishPolicy, args, options);
  if ("report" in fromCore) {
    const report: PolicyPublicationReport | undefined = fromCore.report;
    void report;
    // @ts-expect-error The native policy report is not a string.
    const invalid: string | undefined = fromCore.report;
    void invalid;
  }
  if ("report" in fromClient) {
    const report: PolicyPublicationReport | undefined = fromClient.report;
    void report;
  }
  void client.executeAction(
    ActionNames.PublishPolicy,
    // @ts-expect-error Policy publication requires callerId.
    { titleId: "U_title", policyPath: "policy.json" },
    options
  );
  // @ts-expect-error Acquisition args cannot be passed to policy publication.
  void client.executeAction(ActionNames.PublishPolicy, { appPackagePath: "app.zip" }, options);
}

declare class ReportDriver implements StepDriver {
  execute(
    args: unknown,
    context: DriverContext
  ): Promise<
    ExecutionResult & { report: { operationId: string; outcome: "published" | "failed" } }
  >;
}

export function checkNativeResultInference(
  result: DriverExecutionResult<ReportDriver> | ActionExecutionFailure
): void {
  if ("report" in result) {
    const operationId: string = result.report.operationId;
    const outcome: "published" | "failed" = result.report.outcome;
    // @ts-expect-error Native report fields retain their original types.
    const invalid: number = result.report.operationId;
    void [operationId, outcome, invalid];
  } else {
    const failed: ActionExecutionFailure = result;
    void failed;
    // @ts-expect-error Entry failures do not promise a domain report.
    result.report;
  }
}

export async function checkPublicResultInference(
  client: IFxCoreClient,
  core: FxCore
): Promise<void> {
  const options = { projectPath: "C:\\project" };
  const fromClient: ActionResult<typeof ActionNames.Acquire> = await client.executeAction(
    ActionNames.Acquire,
    { appPackagePath: "app.zip" },
    options
  );
  const fromCore: ActionResult<typeof ActionNames.Acquire> = await core.executeAction(
    ActionNames.Acquire,
    { appPackagePath: "app.zip" },
    options
  );
  const compatible: ExecutionResult = fromClient;
  void [fromCore, compatible];
  // @ts-expect-error Acquisition does not have a domain report.
  fromClient.report;
}

export function checkActionTypes(client: IFxCoreClient, core: FxCore): void {
  const options = { projectPath: "C:\\project" };
  const install: ActionArgs<"teamsApp/extendToM365"> = { appPackagePath: "app.zip" };
  const namedInstall: ActionArgs<typeof ActionNames.Acquire> = install;
  void client.executeAction(ActionNames.Acquire, namedInstall, options);
  // @ts-expect-error Exported action names cannot be reassigned.
  ActionNames.Acquire = ActionNames.Acquire;
  const driverInstall: AcquireArgs = install;
  const publicInstall: PublicAcquireArgs = driverInstall;
  void client.executeAction("teamsApp/extendToM365", publicInstall, options);
  void client.executeAction("teamsApp/extendToM365", install, options);
  void client.executeAction(
    "teamsApp/extendToM365",
    { appPackagePath: "app.zip", scope: "Shared" },
    options
  );
  void core.executeAction("teamsApp/extendToM365", install, options);
  // @ts-expect-error Uninstall stays on the existing API, not the action allowlist.
  void client.executeAction("teamsApp/uninstallFromM365", { titleId: "U_test" }, options);
  // @ts-expect-error Unknown actions are not a public execution surface.
  void client.executeAction("script", {}, options);
  // @ts-expect-error Installation requires an app package path.
  void client.executeAction("teamsApp/extendToM365", {}, options);
  // @ts-expect-error An app package path is not a number.
  void client.executeAction("teamsApp/extendToM365", { appPackagePath: 123 }, options);
  // @ts-expect-error The argument shape must match the selected action.
  void client.executeAction("teamsApp/extendToM365", { titleId: "U_test" }, options);
  void client.executeAction(
    "teamsApp/extendToM365",
    {
      appPackagePath: "app.zip",
      // @ts-expect-error Scope is a closed enumeration.
      scope: "invalid",
    },
    options
  );
  // @ts-expect-error Project context is explicit.
  void client.executeAction("teamsApp/extendToM365", install);
}
