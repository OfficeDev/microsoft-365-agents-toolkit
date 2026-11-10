// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { FxError, Result, err, ok } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { merge } from "lodash";
import { Service } from "typedi";
import isUUID from "validator/lib/isUUID";
import {
  isUsingNewDeveloperPortalApis,
  teamsDevPortalClient,
} from "../../../client/teamsDevPortalClientProvider";
import { AppStudioScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { FileNotFoundError, InvalidActionInputError } from "../../../error/common";
import { getAbsolutePath } from "../../utils/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { AppStudioError } from "./errors";
import { teamsAppUpdateOutputNames } from "./constants";
import { ConfigureTeamsAppArgs } from "./interfaces/ConfigureTeamsAppArgs";
import { AppStudioResultFactory } from "./results";
import { manifestUtils } from "./utils/ManifestUtils";
import { FeatureFlagName } from "../../../common/featureFlags";
import { SovereignCloudEnvironment } from "../../../common/accountUtils";
import {
  appIdentityError,
  getAppTenantId,
  resolveProjectAppResourceId,
} from "../../utils/appIdentity";

export const actionName = "teamsApp/update";

export const internalOutputNames = teamsAppUpdateOutputNames;

@Service(actionName)
export class ConfigureTeamsAppDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.updateDriver");
  readonly progressTitle = getLocalizedString(
    "driver.teamsApp.progressBar.updateTeamsAppStepMessage"
  );

  public async execute(
    args: ConfigureTeamsAppArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.update(args, wrapContext, outputEnvVarNames);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  async update(
    args: ConfigureTeamsAppArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    if (
      process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.GCCH ||
      process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.DOD
    ) {
      context.logProvider.warning(
        getLocalizedString("driver.teamsApp.warning.unsupportedCloud", actionName)
      );
      return ok(new Map<string, string>());
    }

    const result = this.validateArgs(args);
    if (result.isErr()) {
      return err(result.error);
    }

    const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes(),
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }
    const appStudioToken = appStudioTokenRes.value;
    const appPackagePath = getAbsolutePath(args.appPackagePath, context.projectPath);
    if (!(await fs.pathExists(appPackagePath))) {
      return err(
        new FileNotFoundError(
          actionName,
          appPackagePath,
          "https://aka.ms/teamsfx-actions/teamsapp-update"
        )
      );
    }
    const archivedFile = await fs.readFile(appPackagePath);

    // Add capabilities to telemetry properties
    const manifest = manifestUtils.extractManifestFromArchivedFile(archivedFile);
    if (manifest.isErr()) {
      return err(manifest.error);
    }

    const manifestTelemetries = manifestUtils.parseCommonTelemetryProperties(manifest.value);
    merge(context.telemetryProperties, manifestTelemetries);

    // Fail if Teams app not exists, as this action only update the Teams app, not create
    // See work item 17187087
    const manifestId = manifest.value.id;
    if (!isUUID(manifestId)) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.InvalidTeamsAppIdError.name,
          AppStudioError.InvalidTeamsAppIdError.message(manifestId),
          "https://aka.ms/teamsfx-actions/teamsapp-update"
        )
      );
    }
    const identity = await resolveProjectAppResourceId(
      context.projectPath,
      process.env,
      manifestId
    );
    if (identity.isErr()) return err(identity.error);
    const appId = identity.value;
    if (!appId) return err(appIdentityError("APP_ID"));
    if (isUsingNewDeveloperPortalApis() && manifestId !== (process.env.APP_MANIFEST_ID || appId)) {
      return err(appIdentityError("manifest.id"));
    }
    try {
      const appDefinition = await teamsDevPortalClient.getApp(appStudioToken, appId);
      const remoteManifestId =
        "appProfile" in appDefinition
          ? appDefinition.appProfile?.appDetails?.applicationManifest?.id
          : appDefinition.appId;
      if (isUsingNewDeveloperPortalApis() && remoteManifestId && remoteManifestId !== manifestId) {
        return err(appIdentityError("manifest.id"));
      }
    } catch (error) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.TeamsAppNotExistsError.name,
          AppStudioError.TeamsAppNotExistsError.message(appId),
          "https://aka.ms/teamsfx-actions/teamsapp-update"
        )
      );
    }

    try {
      let message = getLocalizedString("driver.teamsApp.progressBar.updateTeamsAppStepMessage");

      const appDefinition = isUsingNewDeveloperPortalApis()
        ? await teamsDevPortalClient.updateApp(appStudioToken, appId, archivedFile)
        : await teamsDevPortalClient.importApp(appStudioToken, archivedFile, true);
      message = getLocalizedString("plugins.appstudio.teamsAppUpdatedLog", appId);
      context.logProvider.verbose(message);
      context.addSummary(message);
      let tenantId: string | undefined;
      let updatedAt: string | undefined;
      if ("appProfile" in appDefinition) {
        const tenant = await getAppTenantId(context.m365TokenProvider);
        if (tenant.isErr()) return err(tenant.error);
        tenantId = tenant.value;
        updatedAt = await teamsDevPortalClient.getAppLastModified(appStudioToken, appId);
      } else {
        tenantId = appDefinition.tenantId;
        updatedAt = appDefinition.updatedAt;
      }
      if (!tenantId || !updatedAt) return err(appIdentityError("tenantId / updatedAt"));
      return ok(
        new Map([
          [internalOutputNames.teamsAppTenantId, tenantId],
          [internalOutputNames.teamsAppUpdateTime, updatedAt],
        ])
      );
    } catch (e: any) {
      return err(
        AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppUpdateFailedError.name,
          AppStudioError.TeamsAppUpdateFailedError.message(appId, e),
          "https://aka.ms/teamsfx-actions/teamsapp-update"
        )
      );
    }
  }

  private validateArgs(args: ConfigureTeamsAppArgs): Result<any, FxError> {
    const invalidParams: string[] = [];
    if (!args || !args.appPackagePath) {
      invalidParams.push("appPackagePath");
    }
    if (invalidParams.length > 0) {
      return err(
        new InvalidActionInputError(
          actionName,
          invalidParams,
          "https://aka.ms/teamsfx-actions/teamsapp-update"
        )
      );
    } else {
      return ok(undefined);
    }
  }
}
