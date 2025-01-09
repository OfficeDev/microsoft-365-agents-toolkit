// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { hooks } from "@feathersjs/hooks/lib";
import { SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import { Service } from "typedi";
import { getLocalizedString } from "../../../common/localizeUtils";
import {
  HttpClientError,
  HttpServerError,
  InvalidActionInputError,
  assembleError,
} from "../../../error/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { UpdateAadAppArgs } from "./interface/updateAadAppArgs";
import { UpdateAadAppOutput } from "./interface/updateAadAppOutput";
import { AadAppClient } from "./utility/aadAppClient";
import { buildAadManifest } from "./utility/buildAadManifest";
import { descriptionMessageKeys, logMessageKeys } from "./utility/constants";
import { AadManifestHelper } from "./utility/aadManifestHelper";
import { AADApplication } from "./interface/AADApplication";
import { AADManifest } from "./interface/AADManifest";
import path from "path";

export const actionName = "aadApp/update"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/aadapp-update";
// logic from src\component\resource\aadApp\aadAppManifestManager.ts
@Service(actionName) // DO NOT MODIFY the service name
export class UpdateAadAppDriver implements StepDriver {
  description = getLocalizedString(descriptionMessageKeys.update);
  readonly progressTitle = getLocalizedString("driver.aadApp.progressBar.updateAadAppTitle");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(args: UpdateAadAppArgs, context: DriverContext): Promise<ExecutionResult> {
    const summaries: string[] = [];

    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));
      const state = this.loadCurrentState();

      this.validateArgs(args);
      const aadAppClient = new AadAppClient(context.m365TokenProvider, context.logProvider);

      let manifest = await buildAadManifest(context, args.manifestPath, args.outputFilePath, state);

      // MS Graph API does not allow adding new OAuth permissions and pre authorize it within one request
      // So split update Microsoft Entra app to two requests:
      // 1. If there's preAuthorizedApplications, remove it temporary and update Microsoft Entra app to create possible new permission

      if (AadManifestHelper.isNewAADManifestSchema(manifest)) {
        manifest = manifest as AADApplication;
        if (
          manifest.api?.preAuthorizedApplications &&
          manifest.api.preAuthorizedApplications?.length > 0
        ) {
          const preAuthorizedApplications = manifest.api.preAuthorizedApplications;
          manifest.api.preAuthorizedApplications = [];
          await aadAppClient.updateAadApp(manifest);
          manifest.api.preAuthorizedApplications = preAuthorizedApplications;
        }
      } else {
        manifest = manifest as AADManifest;
        if (manifest.preAuthorizedApplications && manifest.preAuthorizedApplications.length > 0) {
          const preAuthorizedApplications = manifest.preAuthorizedApplications;
          manifest.preAuthorizedApplications = [];
          await aadAppClient.updateAadApp(manifest);
          manifest.preAuthorizedApplications = preAuthorizedApplications;
        }
      }

      // 2. Update Microsoft Entra app again with full manifest to set preAuthorizedApplications
      await aadAppClient.updateAadApp(manifest);
      const summary = getLocalizedString(
        logMessageKeys.successUpdateAadAppManifest,
        args.manifestPath,
        manifest.id
      );
      context.logProvider?.info(summary);
      summaries.push(summary);

      context.logProvider?.info(
        getLocalizedString(logMessageKeys.successExecuteDriver, actionName)
      );

      const manifestPath = path.isAbsolute(args.manifestPath)
        ? args.manifestPath
        : path.join(context.projectPath, args.manifestPath);
      void AadManifestHelper.showWarningIfManifestIsOutdated(manifestPath, context.projectPath);

      return {
        result: ok(
          new Map(
            Object.entries(state) // convert each property to Map item
              .filter((item) => item[1] && item[1] !== "") // do not return Map item that is empty
          )
        ),
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
      if (
        axios.isAxiosError(error) &&
        error.response // If no response, treat as unhandled error first to understand the actual problem
      ) {
        const message = JSON.stringify(error.response.data);
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
        );
        if (error.response.status >= 400 && error.response.status < 500) {
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
    } finally {
    }
  }

  private validateArgs(args: UpdateAadAppArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.manifestPath !== "string" || !args.manifestPath) {
      invalidParameters.push("manifestPath");
    }

    if (typeof args.outputFilePath !== "string" || !args.outputFilePath) {
      invalidParameters.push("outputFilePath");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }
  }

  private loadCurrentState(): UpdateAadAppOutput {
    return {
      AAD_APP_ACCESS_AS_USER_PERMISSION_ID: process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID,
    };
  }
}
