// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { M365TokenProvider, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { teamsDevPortalClient } from "../../../client/teamsDevPortalClient";
import { AppStudioScopes, GraphScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError, assembleError } from "../../../error/common";
import { QuestionNames } from "../../../question/constants";
import { QuestionMW } from "../../middleware/questionMW";
import { OutputEnvironmentVariableUndefinedError } from "../error/outputEnvironmentVariableUndefinedError";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import {
  OauthRegistration,
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
} from "../teamsApp/interfaces/OauthRegistration";
import { loadStateFromEnv } from "../util/utils";
import { OauthNameTooLongError } from "./error/oauthNameTooLong";
import { CreateOauthArgs } from "./interface/createOauthArgs";
import { CreateOauthOutputs, OutputKeys } from "./interface/createOauthOutputs";
import { logMessageKeys, maxSecretLength, minSecretLength } from "./utility/constants";
import { OauthInfo, getandValidateOauthInfoFromSpec, validateSecret } from "./utility/utility";
import { OauthIdentityProviderInvalid } from "./error/oauthIdentityProviderInvalid";

const actionName = "oauth/register"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/oauth-register";
const supportedFlows = ["authorizationCode"];

@Service(actionName)
export class CreateOauthDriver implements StepDriver {
  description = getLocalizedString("driver.oauth.description.create");
  readonly progressTitle = getLocalizedString("driver.oauth.title.create");

  @hooks([QuestionMW("oauth", true), addStartAndEndTelemetry(actionName, actionName)])
  public async execute(
    args: CreateOauthArgs,
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

      const state = loadStateFromEnv(outputEnvVarNames) as CreateOauthOutputs;
      const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: AppStudioScopes,
      });
      if (appStudioTokenRes.isErr()) {
        throw appStudioTokenRes.error;
      }
      const appStudioToken = appStudioTokenRes.value;

      if (state && state.configurationId) {
        try {
          await teamsDevPortalClient.getOauthRegistrationById(
            appStudioToken,
            state.configurationId
          );
          context.logProvider?.info(
            getLocalizedString(
              logMessageKeys.skipCreateOauth,
              outputEnvVarNames.get(OutputKeys.configurationId)
            )
          );
        } catch (error) {
          context.logProvider?.warning(
            getLocalizedString(
              logMessageKeys.oauthNotFound,
              outputEnvVarNames.get(OutputKeys.configurationId)
            )
          );
        }
      } else {
        const clientId = process.env[QuestionNames.OauthClientId];
        if (clientId) {
          args.clientId = clientId;
        }

        const clientSecret = process.env[QuestionNames.OauthClientSecret];
        if (clientSecret && !args.isPKCEEnabled && args.identityProvider !== "MicrosoftEntra") {
          args.clientSecret = clientSecret;
        }

        this.validateArgs(args);

        const authInfo = await getandValidateOauthInfoFromSpec(args, context, actionName);

        if (args.identityProvider === "MicrosoftEntra") {
          if (!authInfo.authorizationEndpoint!.includes("microsoftonline")) {
            throw new OauthIdentityProviderInvalid(actionName);
          }
        }

        const oauthRegistration = await this.mapArgsToOauthRegistration(
          context.m365TokenProvider,
          args,
          authInfo
        );

        const oauthRegistrationRes = await teamsDevPortalClient.createOauthRegistration(
          appStudioToken,
          oauthRegistration
        );
        outputs.set(
          outputEnvVarNames.get(OutputKeys.configurationId)!,
          oauthRegistrationRes.configurationRegistrationId.oAuthConfigId
        );
        if (
          args.identityProvider === "MicrosoftEntra" &&
          outputEnvVarNames.get(OutputKeys.applicationIdUri)
        ) {
          outputs.set(
            outputEnvVarNames.get(OutputKeys.applicationIdUri)!,
            oauthRegistrationRes.resourceIdentifierUri
          );
        }

        const summary = getLocalizedString(
          logMessageKeys.successCreateOauth,
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

  private validateArgs(args: CreateOauthArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    }

    if (args.name.length > 128) {
      throw new OauthNameTooLongError(actionName);
    }

    if (typeof args.appId !== "string" || !args.appId) {
      invalidParameters.push("appId");
    }

    if (typeof args.apiSpecPath !== "string" || !args.apiSpecPath) {
      invalidParameters.push("apiSpecPath");
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

    if (typeof args.flow !== "string" || !args.flow || !supportedFlows.includes(args.flow)) {
      invalidParameters.push("flow");
    }

    if (typeof args.clientId !== "string" || !args.clientId) {
      invalidParameters.push("clientId");
    }

    if (args.isPKCEEnabled && typeof args.isPKCEEnabled !== "boolean") {
      invalidParameters.push("isPKCEEnabled");
    }

    if (
      args.identityProvider &&
      (typeof args.identityProvider !== "string" ||
        (args.identityProvider !== "Custom" && args.identityProvider !== "MicrosoftEntra"))
    ) {
      invalidParameters.push("identityProvider");
    }

    const isCustomIdentityProvider = !args.identityProvider || args.identityProvider === "Custom";

    if (!args.isPKCEEnabled && isCustomIdentityProvider) {
      if (args.clientSecret && !validateSecret(args.clientSecret)) {
        invalidParameters.push("clientSecret");
      }
    }

    if (args.refreshUrl && typeof args.refreshUrl !== "string") {
      invalidParameters.push("refreshUrl");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }
  }

  private async mapArgsToOauthRegistration(
    tokenProvider: M365TokenProvider,
    args: CreateOauthArgs,
    authInfo: OauthInfo
  ): Promise<OauthRegistration> {
    const currentUserRes = await tokenProvider.getJsonObject({ scopes: GraphScopes });
    if (currentUserRes.isErr()) {
      throw currentUserRes.error;
    }
    const currentUser = currentUserRes.value;
    const userId = currentUser["oid"] as string;

    const targetAudience = args.targetAudience
      ? (args.targetAudience as OauthRegistrationTargetAudience)
      : OauthRegistrationTargetAudience.AnyTenant;
    const applicableToApps = args.applicableToApps
      ? (args.applicableToApps as OauthRegistrationAppType)
      : OauthRegistrationAppType.AnyApp;

    if (args.identityProvider === "MicrosoftEntra") {
      return {
        description: args.name,
        targetUrlsShouldStartWith: authInfo.domain,
        applicableToApps: applicableToApps,
        m365AppId: applicableToApps === OauthRegistrationAppType.SpecificApp ? args.appId : "",
        targetAudience: targetAudience,
        clientId: args.clientId,
        identityProvider: "MicrosoftEntra",
      } as OauthRegistration;
    }

    return {
      description: args.name,
      targetUrlsShouldStartWith: authInfo.domain,
      applicableToApps: applicableToApps,
      m365AppId: applicableToApps === OauthRegistrationAppType.SpecificApp ? args.appId : "",
      targetAudience: targetAudience,
      clientId: args.clientId,
      clientSecret: args.clientSecret ?? "",
      isPKCEEnabled: !!args.isPKCEEnabled,
      authorizationEndpoint: authInfo.authorizationEndpoint,
      tokenExchangeEndpoint: authInfo.tokenExchangeEndpoint,
      tokenRefreshEndpoint: args.refreshUrl ?? authInfo.tokenRefreshEndpoint,
      scopes: authInfo.scopes,
      identityProvider: "Custom",
      // TODO: add this part back after TDP update
      // manageableByUsers: [
      //   {
      //     userId: userId,
      //     accessType: OauthRegistrationUserAccessType.ReadWrite,
      //   },
      // ],
    } as OauthRegistration;
  }
}
