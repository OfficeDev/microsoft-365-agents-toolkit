// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  ok,
  Result,
  SelectFileConfig,
  SelectFolderConfig,
  UserError,
} from "@microsoft/teamsfx-api";
import path from "path";
import * as util from "util";
import VsCodeLogInstance from "../commonlib/log";
import { showError, wrapError } from "../error/common";
import { ExtensionErrors, ExtensionSource } from "../error/error";
import { TeamsAppMigrationHandler } from "../migration/migrationHandler";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";

export async function migrateTeamsTabAppHandler(): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MigrateTeamsTabAppStart);
  const selection = await VS_CODE_UI.showMessage(
    "warn",
    localize("teamstoolkit.migrateTeamsTabApp.warningMessage"),
    true,
    localize("teamstoolkit.migrateTeamsTabApp.upgrade")
  );
  const userCancelError = new UserError(
    ExtensionSource,
    ExtensionErrors.UserCancel,
    localize("teamstoolkit.common.userCancel")
  );
  if (
    selection.isErr() ||
    selection.value !== localize("teamstoolkit.migrateTeamsTabApp.upgrade")
  ) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsTabApp, userCancelError);
    return ok(null);
  }
  const selectFolderConfig: SelectFolderConfig = {
    name: localize("teamstoolkit.migrateTeamsTabApp.selectFolderConfig.name"),
    title: localize("teamstoolkit.migrateTeamsTabApp.selectFolderConfig.title"),
  };
  const selectFolderResult = await VS_CODE_UI.selectFolder(selectFolderConfig);
  if (selectFolderResult.isErr() || selectFolderResult.value.type !== "success") {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsTabApp, userCancelError);
    return ok(null);
  }
  const tabAppPath = selectFolderResult.value.result as string;

  const progressBar = VS_CODE_UI.createProgressBar(
    localize("teamstoolkit.migrateTeamsTabApp.progressTitle"),
    2
  );
  await progressBar.start();

  const migrationHandler = new TeamsAppMigrationHandler(tabAppPath);
  let result: Result<null, FxError> = ok(null);
  let packageUpdated: Result<boolean, FxError> = ok(true);
  let updateFailedFiles: string[] = [];
  try {
    // Update package.json to use @microsoft/teams-js v2
    await progressBar.next(localize("teamstoolkit.migrateTeamsTabApp.updatingPackageJson"));
    VsCodeLogInstance.info(localize("teamstoolkit.migrateTeamsTabApp.updatingPackageJson"));
    packageUpdated = await migrationHandler.updatePackageJson();
    if (packageUpdated.isErr()) {
      throw packageUpdated.error;
    } else if (!packageUpdated.value) {
      // no change in package.json, show warning.
      const warningMessage = util.format(
        localize("teamstoolkit.migrateTeamsTabApp.updatePackageJsonWarning"),
        path.join(tabAppPath, "package.json")
      );
      VsCodeLogInstance.warning(warningMessage);
      void VS_CODE_UI.showMessage("warn", warningMessage, false, "OK");
    } else {
      // Update codes to use @microsoft/teams-js v2
      await progressBar.next(localize("teamstoolkit.migrateTeamsTabApp.updatingCodes"));
      VsCodeLogInstance.info(localize("teamstoolkit.migrateTeamsTabApp.updatingCodes"));
      const failedFiles = await migrationHandler.updateCodes();
      if (failedFiles.isErr()) {
        throw failedFiles.error;
      } else {
        updateFailedFiles = failedFiles.value;
        if (failedFiles.value.length > 0) {
          VsCodeLogInstance.warning(
            util.format(
              localize("teamstoolkit.migrateTeamsTabApp.updateCodesErrorOutput"),
              failedFiles.value.length,
              failedFiles.value.join(", ")
            )
          );
          void VS_CODE_UI.showMessage(
            "warn",
            util.format(
              localize("teamstoolkit.migrateTeamsTabApp.updateCodesErrorMessage"),
              failedFiles.value.length,
              failedFiles.value[0]
            ),
            false,
            "OK"
          );
        }
      }
    }
  } catch (error) {
    result = wrapError(error as Error);
  }

  if (result.isErr()) {
    await progressBar.end(false);
    void showError(result.error);
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsTabApp, result.error);
  } else {
    await progressBar.end(true);
    if (!packageUpdated.isErr() && packageUpdated.value) {
      void VS_CODE_UI.showMessage(
        "info",
        util.format(localize("teamstoolkit.migrateTeamsTabApp.success"), tabAppPath),
        false
      );
    }
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MigrateTeamsTabApp, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.UpdateFailedFiles]: updateFailedFiles.length.toString(),
    });
  }
  return result;
}

export async function migrateTeamsManifestHandler(): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MigrateTeamsManifestStart);
  const selection = await VS_CODE_UI.showMessage(
    "warn",
    localize("teamstoolkit.migrateTeamsManifest.warningMessage"),
    true,
    localize("teamstoolkit.migrateTeamsManifest.upgrade")
  );
  const userCancelError = new UserError(
    ExtensionSource,
    ExtensionErrors.UserCancel,
    localize("teamstoolkit.common.userCancel")
  );
  if (
    selection.isErr() ||
    selection.value !== localize("teamstoolkit.migrateTeamsManifest.upgrade")
  ) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsManifest, userCancelError);
    return ok(null);
  }
  const selectFileConfig: SelectFileConfig = {
    name: localize("teamstoolkit.migrateTeamsManifest.selectFileConfig.name"),
    title: localize("teamstoolkit.migrateTeamsManifest.selectFileConfig.title"),
  };
  const selectFileResult = await VS_CODE_UI.selectFile(selectFileConfig);
  if (selectFileResult.isErr() || selectFileResult.value.type !== "success") {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsManifest, userCancelError);
    return ok(null);
  }
  const manifestPath = selectFileResult.value.result as string;

  const progressBar = VS_CODE_UI.createProgressBar(
    localize("teamstoolkit.migrateTeamsManifest.progressTitle"),
    1
  );
  await progressBar.start();

  const migrationHandler = new TeamsAppMigrationHandler(manifestPath);
  let result: Result<null, FxError> = ok(null);

  try {
    // Update Teams manifest
    await progressBar.next(localize("teamstoolkit.migrateTeamsManifest.updateManifest"));
    VsCodeLogInstance.info(localize("teamstoolkit.migrateTeamsManifest.updateManifest"));
    result = await migrationHandler.updateManifest();
    if (result.isErr()) {
      throw result.error;
    }
  } catch (error) {
    result = wrapError(error as Error);
  }

  if (result.isErr()) {
    await progressBar.end(false);
    void showError(result.error);
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsManifest, result.error);
  } else {
    await progressBar.end(true);
    void VS_CODE_UI.showMessage(
      "info",
      util.format(localize("teamstoolkit.migrateTeamsManifest.success"), manifestPath),
      false
    );
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MigrateTeamsManifest, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
  }
  return result;
}
