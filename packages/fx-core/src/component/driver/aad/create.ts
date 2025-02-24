// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  FxError,
  M365TokenProvider,
  SystemError,
  UserError,
  err,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import { Service } from "typedi";
import { GraphScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import {
  HttpClientError,
  HttpServerError,
  InvalidActionInputError,
  UserCancelError,
  assembleError,
} from "../../../error/common";
import { OutputEnvironmentVariableUndefinedError } from "../error/outputEnvironmentVariableUndefinedError";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { loadStateFromEnv, mapStateToEnv } from "../util/utils";
import { WrapDriverContext } from "../util/wrapUtil";
import { AadAppNameTooLongError } from "./error/aadAppNameTooLongError";
import { MissingEnvUserError } from "./error/missingEnvError";
import { CreateAadAppArgs } from "./interface/createAadAppArgs";
import { CreateAadAppOutput, OutputKeys } from "./interface/createAadAppOutput";
import { SignInAudience } from "./interface/signInAudience";
import { AadAppClient } from "./utility/aadAppClient";
import {
  constants,
  descriptionMessageKeys,
  logMessageKeys,
  questionKeys,
  telemetryKeys,
} from "./utility/constants";
import { AadSet } from "../../../common/globalVars";
import { MissingServiceManagementReferenceError } from "./error/missingServiceManagamentReferenceError";
import { isTestToolEnabledProject } from "../../../common/tools";

const actionName = "aadApp/create"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/aadapp-create";
const driverConstants = {
  generateSecretErrorMessageKey: "driver.aadApp.error.generateSecretFailed",
};

@Service(actionName) // DO NOT MODIFY the service name
export class CreateAadAppDriver implements StepDriver {
  description = getLocalizedString(descriptionMessageKeys.create);
  readonly progressTitle = getLocalizedString("driver.aadApp.progressBar.createAadAppTitle");

  public async execute(
    args: CreateAadAppArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapDriverContext = new WrapDriverContext(context, actionName, actionName);
    return await this.executeInternal(args, wrapDriverContext, outputEnvVarNames);
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  private async executeInternal(
    args: CreateAadAppArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const summaries: string[] = [];
    let outputs: Map<string, string> = new Map<string, string>();
    if (!outputEnvVarNames) {
      const error = new OutputEnvironmentVariableUndefinedError(actionName);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failExecuteDriver, actionName, error.displayMessage)
      );
      return {
        result: err(error),
        summaries: summaries,
      };
    }
    const aadAppState: CreateAadAppOutput = loadStateFromEnv(outputEnvVarNames);
    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));

      this.validateArgs(args);

      const aadAppClient = new AadAppClient(context.m365TokenProvider, context.logProvider);
      if (!aadAppState.clientId) {
        context.logProvider?.info(
          getLocalizedString(
            logMessageKeys.startCreateAadApp,
            outputEnvVarNames.get(OutputKeys.clientId)
          )
        );
        context.addTelemetryProperties({ [telemetryKeys.newAadApp]: "true" });

        const tokenJson = await context.m365TokenProvider.getJsonObject({ scopes: GraphScopes });
        const isMsftAccount =
          tokenJson.isOk() &&
          tokenJson.value.unique_name &&
          (tokenJson.value.unique_name as string).endsWith("@microsoft.com");

        // Create new Microsoft Entra app if no client id exists
        const signInAudience = args.signInAudience
          ? args.signInAudience
          : SignInAudience.AzureADMyOrg;

        // This hidden environment variable is for internal use only.
        const serviceManagementReference =
          args.serviceManagementReference || process.env.TTK_DEFAULT_SERVICE_MANAGEMENT_REFERENCE;

        if (isMsftAccount && !serviceManagementReference) {
          throw new MissingServiceManagementReferenceError(actionName);
        }

        const aadApp = await aadAppClient.createAadApp(
          args.name,
          signInAudience,
          serviceManagementReference
        );
        aadAppState.clientId = aadApp.appId!;
        aadAppState.objectId = aadApp.id!;
        AadSet.add(aadApp.appId!);
        await this.setAadEndpointInfo(context.m365TokenProvider, aadAppState);
        outputs = mapStateToEnv(aadAppState, outputEnvVarNames, [OutputKeys.clientSecret]);

        let summary = getLocalizedString(logMessageKeys.successCreateAadApp, aadApp.id);
        if (isMsftAccount) {
          summary += getLocalizedString(logMessageKeys.deleteAadAfterDebugging);
        }
        context.logProvider?.info(summary);
        summaries.push(summary);
      } else {
        context.logProvider?.info(
          getLocalizedString(
            logMessageKeys.skipCreateAadApp,
            outputEnvVarNames.get(OutputKeys.clientId)
          )
        );
        context.addTelemetryProperties({ [telemetryKeys.newAadApp]: "false" });
      }

      if (args.generateClientSecret) {
        if (!aadAppState.clientSecret) {
          context.logProvider?.info(
            getLocalizedString(
              logMessageKeys.startGenerateClientSecret,
              outputEnvVarNames.get(OutputKeys.clientSecret)
            )
          );
          // Create new client secret if no client secret exists
          if (!aadAppState.objectId) {
            throw new MissingEnvUserError(
              actionName,
              outputEnvVarNames.get(OutputKeys.objectId)!,
              helpLink,
              driverConstants.generateSecretErrorMessageKey
            );
          }

          const clientSecretExpireDays = args.clientSecretExpireDays ?? 180; // Recommended lifetime from Azure Portal
          const clientSecretDescription = args.clientSecretDescription ?? "default";
          aadAppState.clientSecret = await aadAppClient.generateClientSecret(
            aadAppState.objectId,
            clientSecretExpireDays,
            clientSecretDescription
          );
          outputs.set(outputEnvVarNames.get(OutputKeys.clientSecret)!, aadAppState.clientSecret);

          const summary = getLocalizedString(
            logMessageKeys.successGenerateClientSecret,
            aadAppState.objectId
          );
          context.logProvider?.info(summary);
          summaries.push(summary);
        } else {
          context.logProvider?.info(
            getLocalizedString(
              logMessageKeys.skipCreateAadApp,
              outputEnvVarNames.get(OutputKeys.clientSecret)
            )
          );
        }
      }

      context.logProvider?.info(
        getLocalizedString(logMessageKeys.successExecuteDriver, actionName)
      );

      return {
        result: ok(outputs),
        summaries: summaries,
      };
    } catch (error) {
      if (error instanceof UserError || error instanceof SystemError) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, error.displayMessage)
        );
        return {
          result: err(error),
          summaries: summaries,
        };
      }

      if (axios.isAxiosError(error)) {
        const message = JSON.stringify(error.response!.data);
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
        );
        if (error.response!.status >= 400 && error.response!.status < 500) {
          // When user don't have permission to create AAD app, and cannot use Test Tool, we will ask for AAD app id and secret
          if (
            error.response!.status === 403 &&
            message.includes(constants.insufficientPermissionErrorMessage) &&
            !isTestToolEnabledProject(context.projectPath)
          ) {
            context.addTelemetryProperties({
              [telemetryKeys.insufficientPermissionAadApp]: "true",
            });
            const res = await this.askForAADAppIdAndSecret(context, aadAppState, outputEnvVarNames);
            if (res.isOk()) {
              context.addTelemetryProperties({ [telemetryKeys.userInputAadApp]: "true" });
              return {
                result: ok(res.value),
                summaries: summaries,
              };
            } else {
              context.addTelemetryProperties({ [telemetryKeys.userInputAadApp]: "false" });
              return {
                result: err(res.error),
                summaries: summaries,
              };
            }
          }
          return {
            result: err(new HttpClientError(error, actionName, message, helpLink)),
            summaries: summaries,
          };
        } else {
          return {
            result: err(new HttpServerError(error, actionName, message)),
            summaries: summaries,
          };
        }
      }

      const message = JSON.stringify(error);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
      );
      return {
        result: err(assembleError(error as Error, actionName)),
        summaries: summaries,
      };
    }
  }

  /**
   * Pop up a dialog to ask for AAD app id and secret
   * @param context
   * @param aadAppState
   * @param outputEnvVarNames
   * @returns
   */
  async askForAADAppIdAndSecret(
    context: WrapDriverContext,
    aadAppState: CreateAadAppOutput,
    outputEnvVarNames: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    const res = await context.ui!.showMessage(
      "error",
      getLocalizedString(logMessageKeys.insufficientPermission),
      true,
      "Proceed"
    );
    if (res.isOk() && res.value == "Proceed") {
      const aadAppId = await context.ui!.inputText({
        name: "aadAppId",
        title: getLocalizedString(questionKeys.aadAppIdTitle),
        validation: (input: string): string | undefined => {
          if (input.length < 1) {
            return getLocalizedString(questionKeys.addAppIdValidation);
          }
        },
      });
      if (aadAppId.isErr()) {
        return err(new UserCancelError(actionName));
      }
      const aadAppSecret = await context.ui!.inputText({
        name: "aadAppSecret",
        title: getLocalizedString(questionKeys.aadAppSecretTitle),
        password: true,
        validation: (input: string): string | undefined => {
          if (input.length < 1) {
            return getLocalizedString(questionKeys.aadAppSecretValidation);
          }
        },
      });
      if (aadAppSecret.isErr()) {
        return err(new UserCancelError(actionName));
      }
      const aadAppObjectId = await context.ui!.inputText({
        name: "aadAppObjectId",
        title: getLocalizedString(questionKeys.aadAppObjectIdTitle),
        validation: (input: string): string | undefined => {
          if (input.length < 1) {
            return getLocalizedString(questionKeys.aadAppObjectIdValidation);
          }
        },
      });
      if (aadAppObjectId.isErr()) {
        return err(new UserCancelError(actionName));
      }
      aadAppState.clientId = aadAppId.value.result;
      AadSet.add(aadAppState.clientId!);
      aadAppState.clientSecret = aadAppSecret.value.result;
      aadAppState.objectId = aadAppObjectId.value.result;
      const outputs = mapStateToEnv(aadAppState, outputEnvVarNames);
      return ok(outputs);
    } else {
      return err(new UserCancelError(actionName));
    }
  }

  private validateArgs(args: CreateAadAppArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    }

    if (args.generateClientSecret === undefined || typeof args.generateClientSecret !== "boolean") {
      invalidParameters.push("generateClientSecret");
    }

    // Throw error if unexpected signInAudience
    if (
      args.signInAudience &&
      (typeof args.signInAudience !== "string" ||
        !Object.values(SignInAudience).includes(args.signInAudience))
    ) {
      invalidParameters.push("signInAudience");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }

    if (args.name.length > 120) {
      throw new AadAppNameTooLongError(actionName);
    }
  }

  // logic from
  // src\component\resource\aadApp\utils\tokenProvider.ts
  // src\component\resource\aadApp\utils\configs.ts
  private async setAadEndpointInfo(tokenProvider: M365TokenProvider, state: CreateAadAppOutput) {
    const tokenObjectResponse = await tokenProvider.getJsonObject({ scopes: GraphScopes });
    if (tokenObjectResponse.isErr()) {
      throw tokenObjectResponse.error;
    }

    const tenantId = tokenObjectResponse.value.tid as string; // The tid claim is AAD tenant id
    state.tenantId = tenantId;
    state.authorityHost = constants.oauthAuthorityPrefix;
    state.authority = `${constants.oauthAuthorityPrefix}/${tenantId}`;
  }
}
