// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { FxError, Result, err, ok, TeamsAppManifest } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import axios from "axios";
import fs from "fs-extra";
import * as path from "path";
import AdmZip from "adm-zip";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { InstallAppArgs } from "./interfaces/InstallAppArgs";
import { getLocalizedString } from "../../../common/localizeUtils";
import { GraphClient } from "../../../client/graphClient";
import { HttpClientError } from "../../../error/common";
import { InvalidActionInputError, FileNotFoundError } from "../../../error/common";
import { Constants } from "../teamsApp/constants";

const actionName = "devChannel/installApp";

@Service(actionName)
export class InstallAppToChannelDriver implements StepDriver {
  description = getLocalizedString("driver.devChannel.install.description");
  readonly progressTitle = getLocalizedString("driver.devChannel.install.progress.message");

  public async execute(
    args: InstallAppArgs,
    context: DriverContext,
    outputEnvVarNames: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.install(args, wrapContext, outputEnvVarNames);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  async install(
    args: InstallAppArgs,
    context: WrapDriverContext,
    outputEnvVarNames: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    // Need teamId and channelId to install app to channel
    const teamId = process.env["TEAM_ID"];
    const channelId = process.env["CHANNEL_ID"];
    if (!teamId) {
      return err(new InvalidActionInputError(actionName, ["teamId"]));
    }
    if (!channelId) {
      return err(new InvalidActionInputError(actionName, ["channelId"]));
    }

    let appPackagePath = args.appPackagePath;
    if (!path.isAbsolute(appPackagePath)) {
      appPackagePath = path.join(context.projectPath, appPackagePath);
    }
    if (!(await fs.pathExists(appPackagePath))) {
      return err(new FileNotFoundError(actionName, appPackagePath));
    }
    const archivedFile = await fs.readFile(appPackagePath);

    // Read Teams app id from app package.
    const zipEntries = new AdmZip(archivedFile).getEntries();
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
    const teamsAppId = manifest.id;

    try {
      const graphClient = new GraphClient(context.m365TokenProvider);

      // Get installed apps, delete it if externalId already exists.
      const apps = await graphClient.GetAppInstallationForTeam(teamId);
      apps.map(async (app) => {
        if (app.teamsApp.externalId == teamsAppId) {
          await graphClient.DeleteInstalledApp(teamId, app.id);
        }
      });

      await graphClient.InstallAppToChannelAsync(teamId, channelId, archivedFile);
      const message = getLocalizedString("driver.devChannel.install.success", teamId, channelId);
      context.logProvider.info(message);
      context.addSummary(message);
      return ok(new Map<string, string>());
    } catch (error: any) {
      if (axios.isAxiosError(error)) {
        const message = JSON.stringify(error.response!.data);
        context.logProvider.error(message);
        return err(new HttpClientError(error, actionName, message));
      } else {
        return err(error);
      }
    }
  }
}
