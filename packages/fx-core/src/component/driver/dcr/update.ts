// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { FxError, Result, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { teamsGraphClient } from "../../../client/teamsGraphClient";
import { TeamsGraphScopes } from "../../../common/constants";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { assembleError } from "../../../error/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { UpdateDcrArgs } from "./interface/updateDcrArgs";
import { logMessageKeys } from "./utility/constants";
import { isDcrSnapshot } from "./utility/dcrSnapshot";
import { createDcrUpdatePlan, validateUpdateDcrArgs } from "./utility/updateDcr";

const actionName = "dcr/update";

@Service(actionName)
export class UpdateDcrDriver implements StepDriver {
  description = getLocalizedString("driver.dcr.description.update");
  readonly progressTitle = getLocalizedString("driver.dcr.title.update");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(args: unknown, context: DriverContext): Promise<ExecutionResult> {
    context.logProvider.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));
    try {
      const validation = validateUpdateDcrArgs(args);
      if (validation.isErr()) {
        return this.failure(validation.error, context);
      }
      const result = await this.update(validation.value, context);
      if (result.isErr()) {
        return this.failure(result.error, context);
      }
      context.logProvider.info(result.value);
      return { result: ok(new Map()), summaries: [result.value] };
    } catch (error) {
      const failure =
        error instanceof UserError || error instanceof SystemError
          ? error
          : assembleError(error instanceof Error ? error : new Error(String(error)), actionName);
      return this.failure(failure, context);
    }
  }

  private async update(
    args: UpdateDcrArgs,
    context: DriverContext
  ): Promise<Result<string, FxError>> {
    const token = await context.m365TokenProvider.getAccessToken({ scopes: TeamsGraphScopes() });
    if (token.isErr()) {
      return err(token.error);
    }
    const current = await teamsGraphClient.getDcrRegistrationById(
      token.value,
      args.configurationId
    );
    if (!isDcrSnapshot(current, args.configurationId)) {
      return err(
        new SystemError({
          source: actionName,
          name: "InvalidDcrConfiguration",
          message: getDefaultString("driver.dcr.error.invalidConfiguration"),
          displayMessage: getLocalizedString("driver.dcr.error.invalidConfiguration"),
        })
      );
    }
    const plan = createDcrUpdatePlan(args, current);
    if (plan.isErr()) {
      return err(plan.error);
    }
    const changes = {
      ...plan.value.changes,
      ...(plan.value.clearAppBinding ? { m365AppId: "" } : {}),
    };
    if (Object.keys(changes).length === 0) {
      return ok(getLocalizedString("driver.dcr.log.skipUpdate", args.configurationId));
    }

    const inputNames: Record<string, string> = {
      m365AppId: "appId",
      supportedAccountTypes: "includePersonalMicrosoftAccounts",
    };
    const fields = Object.keys(changes)
      .map((field) => inputNames[field] ?? field)
      .join(", ");
    await teamsGraphClient.updateDcrRegistration(token.value, changes, args.configurationId);
    const summary = getLocalizedString(
      "driver.dcr.log.successUpdate",
      args.configurationId,
      fields
    );
    return ok(summary);
  }

  private failure(error: FxError, context: DriverContext): ExecutionResult {
    context.logProvider.error(
      getLocalizedString(
        logMessageKeys.failedExecuteDriver,
        actionName,
        error instanceof UserError || error instanceof SystemError
          ? (error.displayMessage ?? error.message)
          : error.message
      )
    );
    return {
      result: err(error),
      summaries: [],
    };
  }
}
