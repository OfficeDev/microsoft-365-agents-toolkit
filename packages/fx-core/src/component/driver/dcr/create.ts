// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { teamsGraphClient } from "../../../client/teamsGraphClient";
import { TeamsGraphScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError, assembleError } from "../../../error/common";
import { OutputEnvironmentVariableUndefinedError } from "../error/outputEnvironmentVariableUndefinedError";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import {
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
} from "../teamsApp/interfaces/OauthRegistration";
import { DcrRegistration } from "../teamsApp/interfaces/DcrRegistration";
import { loadStateFromEnv } from "../util/utils";
import { DcrNameTooLongError } from "./error/dcrNameTooLong";
import { DcrWellKnownInvalidError } from "./error/dcrWellKnownInvalid";
import { CreateDcrArgs } from "./interface/createDcrArgs";
import { CreateDcrOutputs, OutputKeys } from "./interface/createDcrOutputs";
import { logMessageKeys } from "./utility/constants";
import { validateUrl } from "../oauth/utility/utility";
import {
  mapPersonalMicrosoftAccountsOption,
  validateSupportedAccountTypesCloud,
} from "../oauth/utility/supportedAccountTypes";

const actionName = "dcr/register"; // DO NOT MODIFY the name

@Service(actionName)
export class CreateDcrDriver implements StepDriver {
  description = getLocalizedString("driver.dcr.description.create");
  readonly progressTitle = getLocalizedString("driver.dcr.title.create");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(
    args: CreateDcrArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const summaries: string[] = [];
    const outputs: Map<string, string> = new Map<string, string>();

    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));

      if (!outputEnvVarNames) {
        throw new OutputEnvironmentVariableUndefinedError(actionName);
      }

      const supportedAccountTypesResult = mapPersonalMicrosoftAccountsOption(
        args.includePersonalMicrosoftAccounts,
        actionName
      );
      if (supportedAccountTypesResult.isErr()) {
        throw supportedAccountTypesResult.error;
      }
      const supportedAccountTypes = supportedAccountTypesResult.value;
      const cloudValidation = validateSupportedAccountTypesCloud(supportedAccountTypes, actionName);
      if (cloudValidation.isErr()) {
        throw cloudValidation.error;
      }

      const state = loadStateFromEnv(outputEnvVarNames) as CreateDcrOutputs;

      if (state && state.configurationId) {
        // TGS exposes no GET for DCR — env-var presence is the only idempotency signal.
        context.logProvider?.info(
          getLocalizedString(
            logMessageKeys.skipCreateDcr,
            outputEnvVarNames.get(OutputKeys.configurationId)
          )
        );
      } else {
        const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
          scopes: TeamsGraphScopes(),
        });
        if (appStudioTokenRes.isErr()) {
          throw appStudioTokenRes.error;
        }
        const appStudioToken = appStudioTokenRes.value;

        this.validateArgs(args);

        const applicableToApps = args.applicableToApps
          ? (args.applicableToApps as OauthRegistrationAppType)
          : OauthRegistrationAppType.AnyApp;

        const targetAudience = args.targetAudience
          ? (args.targetAudience as OauthRegistrationTargetAudience)
          : OauthRegistrationTargetAudience.HomeTenant;

        const dcrRegistration: DcrRegistration = {
          clientName: args.name,
          m365AppId:
            applicableToApps === OauthRegistrationAppType.SpecificApp ? (args.appId ?? "") : "",
          applicableToApps: applicableToApps,
          targetAudience: targetAudience,
          targetUrlsShouldStartWith: args.targetUrlsShouldStartWith,
          ...(args.mcpResourceUrl !== undefined ? { mcpResourceUrl: args.mcpResourceUrl } : {}),
          ...(args.wellKnownAuthorizationServer !== undefined
            ? { wellKnownAuthorizationServer: args.wellKnownAuthorizationServer }
            : {}),
          ...(args.resource !== undefined ? { resource: args.resource } : {}),
          ...(supportedAccountTypes !== undefined ? { supportedAccountTypes } : {}),
        };

        const oauthRegistrationRes = await teamsGraphClient.createDcrRegistration(
          appStudioToken,
          dcrRegistration
        );

        outputs.set(
          outputEnvVarNames.get(OutputKeys.configurationId)!,
          oauthRegistrationRes.configurationRegistrationId.oAuthConfigId
        );

        const summary = getLocalizedString(
          logMessageKeys.successCreateDcr,
          oauthRegistrationRes.configurationRegistrationId.oAuthConfigId
        );
        context.logProvider?.info(summary);
        summaries.push(summary);
      }

      return {
        result: ok(outputs),
        summaries: summaries,
      };
    } catch (error) {
      if (error instanceof UserError || error instanceof SystemError) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failedExecuteDriver, actionName, error.displayMessage)
        );
        return {
          result: err(error),
          summaries: summaries,
        };
      }

      const message = JSON.stringify(error);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failedExecuteDriver, actionName, message)
      );
      return {
        result: err(assembleError(error as Error, actionName)),
        summaries: summaries,
      };
    }
  }

  private validateArgs(args: CreateDcrArgs): void {
    const invalidParameters: string[] = [];

    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    } else if (args.name.length > 128) {
      throw new DcrNameTooLongError(actionName);
    }

    if (
      args.applicableToApps === OauthRegistrationAppType.SpecificApp &&
      (typeof args.appId !== "string" || !args.appId)
    ) {
      invalidParameters.push("appId");
    }

    if (args.mcpResourceUrl === undefined && args.wellKnownAuthorizationServer === undefined) {
      invalidParameters.push("mcpResourceUrl", "wellKnownAuthorizationServer");
    }

    if (
      args.mcpResourceUrl !== undefined &&
      (typeof args.mcpResourceUrl !== "string" || !validateUrl(args.mcpResourceUrl))
    ) {
      invalidParameters.push("mcpResourceUrl");
    }

    if (args.wellKnownAuthorizationServer !== undefined) {
      if (
        typeof args.wellKnownAuthorizationServer !== "string" ||
        !args.wellKnownAuthorizationServer
      ) {
        invalidParameters.push("wellKnownAuthorizationServer");
      } else if (!validateUrl(args.wellKnownAuthorizationServer)) {
        throw new DcrWellKnownInvalidError(actionName);
      }
    }

    if (args.resource !== undefined) {
      try {
        if (
          typeof args.resource !== "string" ||
          args.resource.includes("#") ||
          !new URL(args.resource).protocol
        ) {
          invalidParameters.push("resource");
        }
      } catch {
        invalidParameters.push("resource");
      }
    }

    if (
      args.applicableToApps &&
      args.applicableToApps !== OauthRegistrationAppType.AnyApp &&
      args.applicableToApps !== OauthRegistrationAppType.SpecificApp
    ) {
      invalidParameters.push("applicableToApps");
    }

    if (
      args.targetAudience &&
      args.targetAudience !== OauthRegistrationTargetAudience.AnyTenant &&
      args.targetAudience !== OauthRegistrationTargetAudience.HomeTenant
    ) {
      invalidParameters.push("targetAudience");
    }

    if (
      !Array.isArray(args.targetUrlsShouldStartWith) ||
      args.targetUrlsShouldStartWith.length !== 1 ||
      typeof args.targetUrlsShouldStartWith[0] !== "string" ||
      !validateUrl(args.targetUrlsShouldStartWith[0])
    ) {
      invalidParameters.push("targetUrlsShouldStartWith");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters);
    }
  }
}
