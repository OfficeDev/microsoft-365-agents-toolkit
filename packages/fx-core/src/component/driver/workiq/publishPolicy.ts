// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { workIqPolicyClient } from "../../../client/workIqPolicyClient";
import { getResourceServiceEndpoint, ResourceServiceType } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { maskSecret } from "../../../common/stringUtils";
import { InvalidActionInputError } from "../../../error/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { PackageService } from "../../m365/packageService";
import {
  PolicyPublicationOperation,
  PolicyPublicationReport,
  publishPolicyActionName,
  cancelPolicyPublication,
  createPolicyError,
  createPolicyPublicationFailure,
} from "../../workiq/policyPublication";

export type PolicyPublicationExecutionResult = ExecutionResult & {
  report?: PolicyPublicationReport;
};

@Service(publishPolicyActionName)
export class PublishPolicyDriver implements StepDriver {
  readonly description = getLocalizedString("driver.workiq.description");
  readonly progressTitle = getLocalizedString("driver.workiq.progress");

  @hooks([addStartAndEndTelemetry(publishPolicyActionName, publishPolicyActionName)])
  public async execute(
    args: unknown,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<PolicyPublicationExecutionResult> {
    if (
      outputEnvVarNames &&
      [...outputEnvVarNames].some(
        ([key, value]) => key !== "report" || !/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(value)
      )
    ) {
      const error = new InvalidActionInputError(publishPolicyActionName, [
        "writeToEnvironmentFile",
      ]);
      context.logProvider.error(maskSecret(error.displayMessage ?? error.message));
      return { result: err(error), summaries: [] };
    }
    if (!workIqPolicyClient.endpoint) {
      context.logProvider.warning(getLocalizedString("driver.workiq.mockWarning"));
    }
    const mos = new PackageService(
      getResourceServiceEndpoint(ResourceServiceType.MOS3),
      context.logProvider,
      context.m365TokenProvider
    );
    const operation = new PolicyPublicationOperation({
      mos: { prepare: (request) => mos.prepareWorkIqPolicy(request) },
      workIq: {
        prepare: (request) => workIqPolicyClient.prepare(request, context.m365TokenProvider),
      },
    });
    const result = await operation.execute(args, {
      projectPath: context.projectPath,
      signal: context.signal,
    });
    const failure = result.isErr()
      ? result.error
      : context.signal?.aborted
        ? cancelPolicyPublication(result.value)
        : ["published", "unchanged", "simulated"].includes(result.value.outcome)
          ? undefined
          : createPolicyPublicationFailure(
              createPolicyError(
                "POLICY_PUBLISH_INCOMPLETE",
                "driver.workiq.error.incomplete",
                "system"
              ),
              result.value
            );
    const report = failure ? failure.report : result.isOk() ? result.value : result.error.report;
    const onCancellation = () => cancelPolicyPublication(report);
    const summary = getLocalizedString(
      "driver.workiq.summary",
      report.operationId,
      report.outcome,
      report.mos.state,
      report.workIq.state,
      report.nextAction
    );
    if (failure) {
      failure.displayMessage = `${failure.displayMessage ?? failure.message}\n${summary}`;
      context.logProvider.error(maskSecret(failure.displayMessage));
      context.logProvider.info(summary);
      const summaries =
        report.mos.state === "confirmed" ? [getLocalizedString("driver.workiq.mosConfirmed")] : [];
      return { result: err(failure), summaries, report, onCancellation };
    }
    context.logProvider.info(summary);
    const outputs = new Map<string, string>();
    const reportVariable = outputEnvVarNames?.get("report");
    if (reportVariable) outputs.set(reportVariable, JSON.stringify(report));
    return { result: ok(outputs), summaries: [summary], report, onCancellation };
  }
}
