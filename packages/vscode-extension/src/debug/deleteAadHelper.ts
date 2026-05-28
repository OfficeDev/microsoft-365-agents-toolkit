// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { dotenvUtil } from "@microsoft/teamsfx-core/build/component/utils/envUtil";
import M365TokenInstance from "../commonlib/m365Login";
import * as globalVariables from "../globalVariables";
import * as fs from "fs-extra";
import * as path from "path";
import { AadSet, GraphScopes } from "@microsoft/teamsfx-core";
import axios from "axios";
import { ConvertTokenToJson } from "../commonlib/codeFlowLogin";
import VsCodeLogInstance from "../commonlib/log";
import * as util from "util";
import { localize } from "../utils/localizeUtils";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import { FxError } from "@microsoft/teamsfx-api";

const defaultNotificationLocalFile = ".notification.localstore.json";
export async function deleteAad(): Promise<boolean> {
  try {
    if (globalVariables.deleteAadInProgress) {
      return true;
    }
    globalVariables.setDeleteAadInProgress(true);
    const projectPath = globalVariables.workspaceUri!.fsPath;
    const envFile = path.resolve(projectPath, "env", ".env.local");
    const userFile = path.resolve(projectPath, "env", ".env.local.user");
    if (!fs.existsSync(envFile) || !fs.existsSync(userFile)) {
      return true;
    }
    const envData = dotenvUtil.deserialize(fs.readFileSync(envFile, "utf-8"));
    const userEnvData = dotenvUtil.deserialize(fs.readFileSync(userFile, "utf-8"));
    if (!envData.obj["BOT_ID"] && !envData.obj["AAD_APP_CLIENT_ID"]) {
      return true;
    }
    const accountInfo = M365TokenInstance.getCachedAccountInfo();
    if (accountInfo !== undefined) {
      const tokenRes = await M365TokenInstance.getAccessToken({ scopes: GraphScopes });
      if (tokenRes.isErr()) {
        return true;
      }
      const accountJson = ConvertTokenToJson(tokenRes.value);
      const uniqueName = (accountJson as Record<string, string>)["unique_name"];
      if (!uniqueName || !uniqueName.includes("@microsoft.com")) {
        return true;
      }
      VsCodeLogInstance.info(localize("teamstoolkit.localDebug.startDeletingAadProcess"));
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.StartDeleteAadAfterDebug);

      const aadClient = axios.create({
        baseURL: "https://graph.microsoft.com/v1.0",
      });
      aadClient.interceptors.request.use((config) => {
        config.headers["Authorization"] = `Bearer ${tokenRes.value}`;
        return config;
      });
      const list: string[] = [];
      if (envData.obj["BOT_ID"] != undefined && AadSet.has(envData.obj["BOT_ID"])) {
        AadSet.delete(envData.obj["BOT_ID"]);
        list.push(envData.obj["BOT_ID"]);
        envData.obj["BOT_ID"] = "";
        envData.obj["BOT_OBJECT_ID"] = "";
        userEnvData.obj["SECRET_BOT_PASSWORD"] = "";
      }
      if (
        envData.obj["AAD_APP_CLIENT_ID"] != undefined &&
        AadSet.has(envData.obj["AAD_APP_CLIENT_ID"])
      ) {
        AadSet.delete(envData.obj["AAD_APP_CLIENT_ID"]);
        list.push(envData.obj["AAD_APP_CLIENT_ID"]);
        envData.obj["AAD_APP_CLIENT_ID"] = "";
        envData.obj["AAD_APP_OBJECT_ID"] = "";
        envData.obj["AAD_APP_TENANT_ID"] = "";
        envData.obj["AAD_APP_OAUTH_AUTHORITY"] = "";
        envData.obj["AAD_APP_OAUTH_AUTHORITY_HOST"] = "";
        envData.obj["AAD_APP_ACCESS_AS_USER_PERMISSION_ID"] = "";
        userEnvData.obj["SECRET_AAD_APP_CLIENT_SECRET"] = "";
      }
      VsCodeLogInstance.info(localize("teamstoolkit.localDebug.updatingLocalEnvFile"));
      fs.writeFileSync(envFile, dotenvUtil.serialize(envData));
      fs.writeFileSync(userFile, dotenvUtil.serialize(userEnvData));
      VsCodeLogInstance.info(localize("teamstoolkit.localDebug.successUpdateLocalEnvFile"));
      if (fs.existsSync(path.resolve(projectPath, defaultNotificationLocalFile))) {
        VsCodeLogInstance.info(
          localize("teamstoolkit.localDebug.startDeletingNotificationLocalStoreFile")
        );
        fs.writeFileSync(path.resolve(projectPath, defaultNotificationLocalFile), "{}");
        VsCodeLogInstance.info(
          localize("teamstoolkit.localDebug.successDeleteNotificationLocalStoreFile")
        );
      }
      for (const id of list) {
        try {
          VsCodeLogInstance.info(
            util.format(localize("teamstoolkit.localDebug.startDeletingAadApp"), id)
          );
          await aadClient.delete(`applications(appId='${id}')`);
          VsCodeLogInstance.info(
            util.format(localize("teamstoolkit.localDebug.successDeleteAadApp"), id)
          );
        } catch (error) {
          VsCodeLogInstance.warning(
            util.format(
              localize("teamstoolkit.localDebug.failDeleteAadApp"),
              id,
              (error as Error).toString()
            )
          );
        }
      }
      VsCodeLogInstance.info(localize("teamstoolkit.localDebug.successDeleteAadProcess"));
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.SuccessDeleteAadAfterDebug);
    }
    return true;
  } catch (error) {
    VsCodeLogInstance.warning(
      util.format(
        localize("teamstoolkit.localDebug.failDeleteAadProcess"),
        (error as Error).toString()
      )
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.FailDeleteAadAfterDebug, error as FxError);
    return false;
  } finally {
    globalVariables.setDeleteAadInProgress(false);
  }
}
