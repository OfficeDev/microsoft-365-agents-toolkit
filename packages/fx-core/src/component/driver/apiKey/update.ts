// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { teamsDevPortalClient } from "../../../client/teamsDevPortalClient";
import { AppStudioScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError, assembleError } from "../../../error";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import {
  ApiSecretRegistration,
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationTargetAudience,
  ApiSecretRegistrationUpdate,
} from "../teamsApp/interfaces/ApiSecretRegistration";
import { ApiKeyNameTooLongError } from "./error/apiKeyNameTooLong";
import { UpdateApiKeyArgs } from "./interface/updateApiKeyArgs";
import { logMessageKeys } from "./utility/constants";
import { getDomain, validateDomain } from "./utility/utility";
import { WrapDriverContext } from "../util/wrapUtil";

const actionName = "apiKey/update"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/apiKey-update";

@Service(actionName) // DO NOT MODIFY the service name
export class UpdateApiKeyDriver implements StepDriver {
  description = getLocalizedString("driver.apiKey.description.update");
  readonly progressTitle = getLocalizedString("driver.aadApp.apiKey.title.update");

  public async execute(args: UpdateApiKeyArgs, context: DriverContext): Promise<ExecutionResult> {
    const wrapDriverContext = new WrapDriverContext(context, actionName, actionName);
    const result = await this._run(args, wrapDriverContext);
    return result;
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async _run(args: UpdateApiKeyArgs, context: WrapDriverContext): Promise<ExecutionResult> {
    const summaries: string[] = [];
    const outputs: Map<string, string> = new Map<string, string>();

    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));
      this.validateArgs(args);

      const domain = await getDomain(args, context, actionName);
      validateDomain(domain, actionName);

      const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: AppStudioScopes,
      });
      if (appStudioTokenRes.isErr()) {
        throw appStudioTokenRes.error;
      }
      const appStudioToken = appStudioTokenRes.value;

      const getApiKeyRes = await teamsDevPortalClient.getApiKeyRegistrationById(
        appStudioToken,
        args.registrationId
      );
      const diffMsgs = this.compareApiKeyRegistration(getApiKeyRes, args, domain);
      // If there is no difference, skip the update
      if (!diffMsgs || diffMsgs.length === 0) {
        const summary = getLocalizedString(logMessageKeys.skipUpdateApiKey);
        context.logProvider?.info(summary);
        summaries.push(summary);

        return {
          result: ok(outputs),
          summaries: summaries,
        };
      }

      // If there is difference, ask user to confirm the update
      // Skip confirm if only targetUrlsShouldStartWith is different when the url contains devtunnel
      if (!this.shouldSkipConfirm(diffMsgs, getApiKeyRes.targetUrlsShouldStartWith, domain)) {
        const userConfirm = await context.ui!.confirm!({
          name: "confirm-update-api-key",
          title: getLocalizedString("driver.apiKey.confirm.update", diffMsgs.join(",\n")),
          default: true,
        });
        if (userConfirm.isErr()) {
          throw userConfirm.error;
        }
      }

      const apiKey = this.mapArgsToApiSecretRegistration(args, domain);
      await teamsDevPortalClient.updateApiKeyRegistration(
        appStudioToken,
        apiKey,
        args.registrationId
      );

      void context.ui!.showMessage(
        "info",
        getLocalizedString("driver.apiKey.info.update", diffMsgs.join(",\n")),
        false
      );
      const summary = getLocalizedString(logMessageKeys.successUpdateApiKey);
      context.logProvider?.info(summary);
      summaries.push(summary);

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

  private validateArgs(args: UpdateApiKeyArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.registrationId !== "string" || !args.registrationId) {
      invalidParameters.push("registrationId");
    }

    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    }

    if (args.name.length > 128) {
      throw new ApiKeyNameTooLongError(actionName);
    }

    if (typeof args.appId !== "string" || !args.appId) {
      invalidParameters.push("appId");
    }

    if (typeof args.apiSpecPath !== "string" || !args.apiSpecPath) {
      invalidParameters.push("apiSpecPath");
    }

    if (
      args.applicableToApps &&
      args.applicableToApps !== ApiSecretRegistrationAppType.AnyApp &&
      args.applicableToApps !== ApiSecretRegistrationAppType.SpecificApp
    ) {
      invalidParameters.push("applicableToApps");
    }

    if (
      args.targetAudience &&
      args.targetAudience !== ApiSecretRegistrationTargetAudience.AnyTenant &&
      args.targetAudience !== ApiSecretRegistrationTargetAudience.HomeTenant
    ) {
      invalidParameters.push("targetAudience");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }
  }

  private compareApiKeyRegistration(
    current: ApiSecretRegistration,
    input: UpdateApiKeyArgs,
    domain: string[]
  ): string[] {
    const diffMsgs: string[] = [];
    if (current.description !== input.name) {
      diffMsgs.push(`description: ${current.description as string} => ${input.name}`);
    }

    if (input.applicableToApps && current.applicableToApps !== input.applicableToApps) {
      let msg = `applicableToApps: ${current.applicableToApps} => ${input.applicableToApps}`;
      if (input.applicableToApps === "SpecificApp") {
        msg += `, specificAppId: ${input.appId}`;
      }
      diffMsgs.push(msg);
    }

    if (input.targetAudience && current.targetAudience !== input.targetAudience) {
      diffMsgs.push(
        `targetAudience: ${current.targetAudience as string} => ${input.targetAudience}`
      );
    }

    // Compare domain
    if (
      current.targetUrlsShouldStartWith.length !== domain.length ||
      !current.targetUrlsShouldStartWith.every((value) => domain.includes(value)) ||
      !domain.every((value) => current.targetUrlsShouldStartWith.includes(value))
    ) {
      diffMsgs.push(
        `targetUrlsShouldStartWith: ${current.targetUrlsShouldStartWith.join(",")} => ${domain.join(
          ","
        )}`
      );
    }

    return diffMsgs;
  }

  private mapArgsToApiSecretRegistration(
    args: UpdateApiKeyArgs,
    domain: string[]
  ): ApiSecretRegistrationUpdate {
    const targetAudience = args.targetAudience
      ? (args.targetAudience as ApiSecretRegistrationTargetAudience)
      : undefined;
    const applicableToApps = args.applicableToApps
      ? (args.applicableToApps as ApiSecretRegistrationAppType)
      : undefined;

    const apiKey: ApiSecretRegistrationUpdate = {
      description: args.name,
      targetUrlsShouldStartWith: domain,
      applicableToApps: applicableToApps,
      specificAppId:
        applicableToApps === ApiSecretRegistrationAppType.SpecificApp ? args.appId : "",
      targetAudience: targetAudience,
    };

    return apiKey;
  }

  // Should skip confirm box if only targetUrlsShouldStartWith is different and the url contains devtunnel
  private shouldSkipConfirm(diffMsgs: string[], getDomain: string[], domain: string[]): boolean {
    return (
      diffMsgs.length === 1 &&
      diffMsgs[0].includes("targetUrlsShouldStartWith") &&
      getDomain.length === domain.length &&
      getDomain.every((value) => value.includes("devtunnel")) &&
      domain.every((value) => value.includes("devtunnel"))
    );
  }
}
