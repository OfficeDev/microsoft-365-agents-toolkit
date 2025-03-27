// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { M365TokenProvider, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { teamsDevPortalClient } from "../../../client/teamsDevPortalClient";
import { AppStudioScopes, GraphScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError, assembleError } from "../../../error";
import { QuestionNames } from "../../../question/constants";
import { QuestionMW } from "../../middleware/questionMW";
import { OutputEnvironmentVariableUndefinedError } from "../error/outputEnvironmentVariableUndefinedError";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import {
  ApiSecretRegistration,
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationTargetAudience,
  ApiSecretRegistrationUserAccessType,
} from "../teamsApp/interfaces/ApiSecretRegistration";
import { ApiSecretRegistrationClientSecret } from "../teamsApp/interfaces/ApiSecretRegistrationClientSecret";
import { ApiKeyClientSecretInvalidError } from "./error/apiKeyClientSecretInvalid";
import { ApiKeyNameTooLongError } from "./error/apiKeyNameTooLong";
import { CreateApiKeyArgs } from "./interface/createApiKeyArgs";
import { CreateApiKeyOutputs, OutputKeys } from "./interface/createApiKeyOutputs";
import { logMessageKeys, maxSecretLength, minSecretLength } from "./utility/constants";
import { getDomain, loadStateFromEnv, validateDomain } from "./utility/utility";
import { apiKeyFromScratchClientSecretInvalid } from "./error/apiKeyFromScratchClientSecretInvalid";
import { WrapDriverContext } from "../util/wrapUtil";

const actionName = "apiKey/register"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/apiKey-register";

@Service(actionName) // DO NOT MODIFY the service name
export class CreateApiKeyDriver implements StepDriver {
  description = getLocalizedString("driver.apiKey.description.create");
  readonly progressTitle = getLocalizedString("driver.aadApp.apiKey.title.create");

  public async execute(
    args: CreateApiKeyArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapDriverContext = new WrapDriverContext(context, actionName, actionName);
    const result = await this._run(args, wrapDriverContext, outputEnvVarNames);
    return result;
  }

  @hooks([QuestionMW("apiKey", true), addStartAndEndTelemetry(actionName, actionName)])
  public async _run(
    args: CreateApiKeyArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const summaries: string[] = [];
    const outputs: Map<string, string> = new Map<string, string>();

    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));

      if (!outputEnvVarNames) {
        throw new OutputEnvironmentVariableUndefinedError(actionName);
      }

      const state = loadStateFromEnv(outputEnvVarNames) as CreateApiKeyOutputs;
      const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: AppStudioScopes,
      });
      if (appStudioTokenRes.isErr()) {
        throw appStudioTokenRes.error;
      }
      const appStudioToken = appStudioTokenRes.value;

      if (state && state.registrationId) {
        try {
          await teamsDevPortalClient.getApiKeyRegistrationById(
            appStudioToken,
            state.registrationId
          );
          context.logProvider?.info(
            getLocalizedString(
              logMessageKeys.skipCreateApiKey,
              outputEnvVarNames.get(OutputKeys.registrationId)
            )
          );
        } catch (error) {
          context.logProvider?.warning(
            getLocalizedString(
              logMessageKeys.apiKeyNotFound,
              outputEnvVarNames.get(OutputKeys.registrationId)
            )
          );
        }
      } else {
        const clientSecret = this.loadClientSecret();
        if (clientSecret) {
          args.primaryClientSecret = clientSecret;
        }

        this.validateArgs(args);

        const domains = await getDomain(args, context, actionName);
        validateDomain(domains, actionName);

        const apiKey = await this.mapArgsToApiSecretRegistration(
          context.m365TokenProvider,
          args,
          domains
        );

        const apiRegistrationRes = await teamsDevPortalClient.createApiKeyRegistration(
          appStudioToken,
          apiKey
        );
        outputs.set(outputEnvVarNames.get(OutputKeys.registrationId)!, apiRegistrationRes.id!);

        const summary = getLocalizedString(
          logMessageKeys.successCreateApiKey,
          apiRegistrationRes.id!
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

  private loadClientSecret(): string | undefined {
    const clientSecret = process.env[QuestionNames.ApiSpecApiKey];
    return clientSecret;
  }

  private validateSecret(apiKeySecret: string): boolean {
    if (typeof apiKeySecret !== "string") {
      return false;
    }

    if (apiKeySecret.length > maxSecretLength || apiKeySecret.length < minSecretLength) {
      return false;
    }

    return true;
  }

  private validateArgs(args: CreateApiKeyArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    }

    if (args.name.length > 128) {
      throw new ApiKeyNameTooLongError(actionName);
    }

    if (typeof args.appId !== "string" || !args.appId) {
      invalidParameters.push("appId");
    }

    if (args.primaryClientSecret && !this.validateSecret(args.primaryClientSecret)) {
      throw args.primaryClientSecret === " "
        ? new apiKeyFromScratchClientSecretInvalid(actionName)
        : new ApiKeyClientSecretInvalidError(actionName);
    }

    if (args.secondaryClientSecret && !this.validateSecret(args.secondaryClientSecret)) {
      throw new ApiKeyClientSecretInvalidError(actionName);
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

  private async mapArgsToApiSecretRegistration(
    tokenProvider: M365TokenProvider,
    args: CreateApiKeyArgs,
    domain: string[]
  ): Promise<ApiSecretRegistration> {
    const currentUserRes = await tokenProvider.getJsonObject({ scopes: GraphScopes });
    if (currentUserRes.isErr()) {
      throw currentUserRes.error;
    }
    const currentUser = currentUserRes.value;
    const userId = currentUser["oid"] as string;

    const secrets = [];
    if (args.primaryClientSecret) {
      secrets.push(args.primaryClientSecret);
    }
    if (args.secondaryClientSecret) {
      secrets.push(args.secondaryClientSecret);
    }
    let isPrimary = true;
    const clientSecrets = secrets.map((secret) => {
      const clientSecret: ApiSecretRegistrationClientSecret = {
        value: secret,
        description: args.name,
        priority: isPrimary ? 0 : 1,
        isValueRedacted: false,
      };
      isPrimary = false;
      return clientSecret;
    });

    const targetAudience: ApiSecretRegistrationTargetAudience = args.targetAudience
      ? (args.targetAudience as ApiSecretRegistrationTargetAudience)
      : ApiSecretRegistrationTargetAudience.AnyTenant;
    const applicableToApps: ApiSecretRegistrationAppType = args.applicableToApps
      ? (args.applicableToApps as ApiSecretRegistrationAppType)
      : ApiSecretRegistrationAppType.AnyApp;

    const apiKey: ApiSecretRegistration = {
      description: args.name,
      targetUrlsShouldStartWith: domain,
      applicableToApps: applicableToApps,
      specificAppId:
        applicableToApps === ApiSecretRegistrationAppType.SpecificApp ? args.appId : "",
      targetAudience: targetAudience,
      clientSecrets: clientSecrets,
      manageableByUsers: [
        {
          userId: userId,
          accessType: ApiSecretRegistrationUserAccessType.ReadWrite,
        },
      ],
    };

    return apiKey;
  }
}
