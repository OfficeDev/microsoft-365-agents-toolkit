// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { FxError, Platform, Result, TeamsAppManifest, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import { merge } from "lodash";
import { Service } from "typedi";
import { teamsDevPortalClient } from "../../../client/teamsDevPortalClient";
import { AppStudioScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { FileNotFoundError, InvalidActionInputError, UserCancelError } from "../../../error/common";
import { getAbsolutePath } from "../../utils/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { Constants } from "./constants";
import { PublishAppPackageArgs } from "./interfaces/PublishAppPackageArgs";
import { TelemetryPropertyKey } from "./utils/telemetry";
import { ODRProvider } from "../../utils/odrProvider";
import { exec } from "child_process";
import { promisify } from "util";
import { LocalMcpPrefix } from "../../constants";
import { AppStudioError } from "./errors";
import { AppStudioResultFactory } from "./results";

export const actionName = "teamsApp/publishAppPackage";

const defaultOutputNames = {
  publishedAppId: "TEAMS_APP_PUBLISHED_APP_ID",
};

@Service(actionName)
export class PublishAppPackageDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.publishDriver");
  readonly progressTitle = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.2");

  public async execute(
    args: PublishAppPackageArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.publish(args, wrapContext, outputEnvVarNames);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async publish(
    args: PublishAppPackageArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    const argsValidationResult = this.validateArgs(args);
    if (argsValidationResult.isErr()) {
      return err(argsValidationResult.error);
    }

    if (!outputEnvVarNames) {
      outputEnvVarNames = new Map(Object.entries(defaultOutputNames));
    }

    const appPackagePath = getAbsolutePath(args.appPackagePath, context.projectPath);
    if (!(await fs.pathExists(appPackagePath))) {
      return err(
        new FileNotFoundError(
          actionName,
          appPackagePath,
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    }
    const archivedFile = await fs.readFile(appPackagePath);

    const zipEntries = new AdmZip(archivedFile).getEntries();

    const pluginFile = zipEntries.find((x) => x.entryName === "ai-plugin.json");
    if (pluginFile) {
      const isValid = await this.verifyLocalPluginCerts(pluginFile);

      if (!isValid) {
        const message =
          "MSIX package certificate verification failed. One or more local mcp servers do not have a cert or have a self-signed cert.";
        return err(
          AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
            message,
            message,
          ])
        );
      }
    }

    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
    if (!manifestFile) {
      return err(
        new FileNotFoundError(
          actionName,
          Constants.MANIFEST_FILE,
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    }
    const manifestString = manifestFile.getData().toString();
    const manifest = JSON.parse(manifestString) as TeamsAppManifest;

    // manifest.id === externalID
    const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }

    let result;

    const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep1");
    context.addSummary(message);

    try {
      const existApp = await teamsDevPortalClient.getStaggedApp(
        appStudioTokenRes.value,
        manifest.id
      );
      if (existApp) {
        context.addSummary(
          getLocalizedString("driver.teamsApp.summary.publishTeamsAppExists", manifest.id)
        );
        let executePublishUpdate = false;
        let description = getLocalizedString(
          "plugins.appstudio.pubWarn",
          existApp.displayName,
          existApp.publishingState
        );
        if (existApp.lastModifiedDateTime) {
          description =
            description +
            getLocalizedString(
              "plugins.appstudio.lastModified",
              existApp.lastModifiedDateTime?.toLocaleString()
            );
        }
        description =
          description + getLocalizedString("plugins.appstudio.updatePublihsedAppConfirm");
        const confirm = getLocalizedString("core.option.confirm");
        const res = await context.ui?.showMessage("warn", description, true, confirm);
        if (res?.isOk() && res.value === confirm) executePublishUpdate = true;

        if (executePublishUpdate) {
          const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.1");
          context.addSummary(message);
          context.logProvider.debug(message);
          const appId = await teamsDevPortalClient.publishTeamsAppUpdate(
            appStudioTokenRes.value,
            manifest.id,
            archivedFile
          );
          result = new Map([[outputEnvVarNames.get("publishedAppId") as string, appId]]);
          merge(context.telemetryProperties, {
            [TelemetryPropertyKey.updateExistingApp]: "true",
            [TelemetryPropertyKey.publishedAppId]: appId,
          });
        } else {
          return err(new UserCancelError(actionName));
        }
      } else {
        context.addSummary(
          getLocalizedString("driver.teamsApp.summary.publishTeamsAppNotExists", manifest.id)
        );
        const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.2");
        context.addSummary(message);
        context.logProvider.debug(message);
        const appId = await teamsDevPortalClient.publishTeamsApp(
          appStudioTokenRes.value,
          manifest.id,
          archivedFile
        );
        result = new Map([[outputEnvVarNames.get("publishedAppId") as string, appId]]);
        merge(context.telemetryProperties, {
          [TelemetryPropertyKey.updateExistingApp]: "false",
        });
      }
    } catch (e: any) {
      return err(e);
    }

    context.logProvider.info(`Publish success!`);
    context.addSummary(
      getLocalizedString("driver.teamsApp.summary.publishTeamsAppSuccess", manifest.id)
    );
    if (context.platform === Platform.CLI) {
      const msg = getLocalizedString(
        "plugins.appstudio.publishSucceedNotice.cli",
        manifest.name.short,
        Constants.TEAMS_ADMIN_PORTAL,
        Constants.TEAMS_MANAGE_APP_DOC
      );
      context.ui?.showMessage("info", msg, false);
    }
    return ok(result);
  }

  private validateArgs(args: PublishAppPackageArgs): Result<any, FxError> {
    const invalidParams: string[] = [];
    if (!args || !args.appPackagePath) {
      invalidParams.push("appPackagePath");
    }
    if (invalidParams.length > 0) {
      return err(
        new InvalidActionInputError(
          actionName,
          invalidParams,
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    } else {
      return ok(undefined);
    }
  }

  private async verifyLocalPluginCerts(pluginFile: AdmZip.IZipEntry): Promise<boolean> {
    const pluginContent = pluginFile.getData().toString();
    const pluginContentAsJson = JSON.parse(pluginContent);

    const servers = await ODRProvider.listServers();

    let allValidCerts = true;

    pluginContentAsJson.runtimes
      .filter((runtime: any) => runtime.type === "LocalPlugin")
      .forEach(async (runtime: any) => {
        const serverInfo = servers.find(
          (x) => x.identifier === runtime.spec.local_endpoint.substring(LocalMcpPrefix.length)
        );
        if (serverInfo) {
          const valid = await this.verifyPackageFamilyCertIsValid(serverInfo.packageFamily);

          if (!valid) {
            allValidCerts = false;
          }
        } else {
          allValidCerts = false;
        }
      });

    return allValidCerts;
  }

  private async verifyPackageFamilyCertIsValid(packageName: string): Promise<boolean> {
    const execAsync = promisify(exec);
    const command = `powershell.exe -Command "& Get-AppxPackage | where { $_.PackageFamilyName -eq '${packageName}' } | select { $_.SignatureKind }"`;

    try {
      const { stdout } = await execAsync(command);

      if (!stdout) {
        return false;
      }

      if (stdout.toLowerCase().includes("developer")) {
        return false;
      }
      return true;
    } catch (error) {
      console.error("Unable to get cert info for package name", error);
      return false;
    }
  }
}
