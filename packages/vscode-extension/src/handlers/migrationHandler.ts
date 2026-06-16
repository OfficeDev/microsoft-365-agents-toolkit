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

export const migrationHandlerOps = {
  sendTelemetryEvent: (eventName: string, properties?: any) =>
    ExtTelemetry.sendTelemetryEvent(eventName as any, properties),
  sendTelemetryErrorEvent: (eventName: string, error: FxError) =>
    ExtTelemetry.sendTelemetryErrorEvent(eventName as any, error),
  showMessage: (
    messageLevel: "warn" | "error" | "info",
    message: string,
    modal: boolean,
    ...items: string[]
  ) => VS_CODE_UI.showMessage(messageLevel, message, modal, ...items),
  selectFolder: (config: SelectFolderConfig) => VS_CODE_UI.selectFolder(config),
  selectFile: (config: SelectFileConfig) => VS_CODE_UI.selectFile(config),
  createProgressBar: (title: string, totalSteps: number) =>
    VS_CODE_UI.createProgressBar(title, totalSteps),
  createMigrationHandler: (targetPath: string) => new TeamsAppMigrationHandler(targetPath),
  showError: (error: FxError) => showError(error),
  wrapError: (error: Error) => wrapError(error),
  localize: (key: string, ...args: any[]) => localize(key, ...args),
};
const migrationHandlerDeps = migrationHandlerOps;

export async function migrateTeamsTabAppHandler(): Promise<Result<null, FxError>> {
  migrationHandlerDeps.sendTelemetryEvent(TelemetryEvent.MigrateTeamsTabAppStart);
  const selection = await migrationHandlerDeps.showMessage(
    "warn",
    migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.warningMessage"),
    true,
    migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.upgrade")
  );
  const userCancelError = new UserError(
    ExtensionSource,
    ExtensionErrors.UserCancel,
    migrationHandlerDeps.localize("teamstoolkit.common.userCancel")
  );
  if (
    selection.isErr() ||
    selection.value !== migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.upgrade")
  ) {
    migrationHandlerDeps.sendTelemetryErrorEvent(
      TelemetryEvent.MigrateTeamsTabApp,
      userCancelError
    );
    return ok(null);
  }
  const selectFolderConfig: SelectFolderConfig = {
    name: migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.selectFolderConfig.name"),
    title: migrationHandlerDeps.localize(
      "teamstoolkit.migrateTeamsTabApp.selectFolderConfig.title"
    ),
  };
  const selectFolderResult = await migrationHandlerDeps.selectFolder(selectFolderConfig);
  if (selectFolderResult.isErr() || selectFolderResult.value.type !== "success") {
    migrationHandlerDeps.sendTelemetryErrorEvent(
      TelemetryEvent.MigrateTeamsTabApp,
      userCancelError
    );
    return ok(null);
  }
  const tabAppPath = selectFolderResult.value.result as string;

  const progressBar = migrationHandlerDeps.createProgressBar(
    migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.progressTitle"),
    2
  );
  await progressBar.start();

  const migrationHandler = migrationHandlerDeps.createMigrationHandler(tabAppPath);
  let result: Result<null, FxError> = ok(null);
  let packageUpdated: Result<boolean, FxError> = ok(true);
  let updateFailedFiles: string[] = [];
  try {
    // Update package.json to use @microsoft/teams-js v2
    await progressBar.next(
      migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.updatingPackageJson")
    );
    VsCodeLogInstance.info(
      migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.updatingPackageJson")
    );
    packageUpdated = await migrationHandler.updatePackageJson();
    if (packageUpdated.isErr()) {
      throw packageUpdated.error;
    } else if (!packageUpdated.value) {
      // no change in package.json, show warning.
      const warningMessage = util.format(
        migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.updatePackageJsonWarning"),
        path.join(tabAppPath, "package.json")
      );
      VsCodeLogInstance.warning(warningMessage);
      void migrationHandlerDeps.showMessage("warn", warningMessage, false, "OK");
    } else {
      // Update codes to use @microsoft/teams-js v2
      await progressBar.next(
        migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.updatingCodes")
      );
      VsCodeLogInstance.info(
        migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.updatingCodes")
      );
      const failedFiles = await migrationHandler.updateCodes();
      if (failedFiles.isErr()) {
        throw failedFiles.error;
      } else {
        updateFailedFiles = failedFiles.value;
        if (failedFiles.value.length > 0) {
          VsCodeLogInstance.warning(
            util.format(
              migrationHandlerDeps.localize(
                "teamstoolkit.migrateTeamsTabApp.updateCodesErrorOutput"
              ),
              failedFiles.value.length,
              failedFiles.value.join(", ")
            )
          );
          void migrationHandlerDeps.showMessage(
            "warn",
            util.format(
              migrationHandlerDeps.localize(
                "teamstoolkit.migrateTeamsTabApp.updateCodesErrorMessage"
              ),
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
    result = migrationHandlerDeps.wrapError(error as Error);
  }

  if (result.isErr()) {
    await progressBar.end(false);
    void migrationHandlerDeps.showError(result.error);
    migrationHandlerDeps.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsTabApp, result.error);
  } else {
    await progressBar.end(true);
    if (!packageUpdated.isErr() && packageUpdated.value) {
      void migrationHandlerDeps.showMessage(
        "info",
        util.format(
          migrationHandlerDeps.localize("teamstoolkit.migrateTeamsTabApp.success"),
          tabAppPath
        ),
        false
      );
    }
    migrationHandlerDeps.sendTelemetryEvent(TelemetryEvent.MigrateTeamsTabApp, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.UpdateFailedFiles]: updateFailedFiles.length.toString(),
    });
  }
  return result;
}

export async function migrateTeamsManifestHandler(): Promise<Result<null, FxError>> {
  migrationHandlerDeps.sendTelemetryEvent(TelemetryEvent.MigrateTeamsManifestStart);
  const selection = await migrationHandlerDeps.showMessage(
    "warn",
    migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.warningMessage"),
    true,
    migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.upgrade")
  );
  const userCancelError = new UserError(
    ExtensionSource,
    ExtensionErrors.UserCancel,
    migrationHandlerDeps.localize("teamstoolkit.common.userCancel")
  );
  if (
    selection.isErr() ||
    selection.value !== migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.upgrade")
  ) {
    migrationHandlerDeps.sendTelemetryErrorEvent(
      TelemetryEvent.MigrateTeamsManifest,
      userCancelError
    );
    return ok(null);
  }
  const selectFileConfig: SelectFileConfig = {
    name: migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.selectFileConfig.name"),
    title: migrationHandlerDeps.localize(
      "teamstoolkit.migrateTeamsManifest.selectFileConfig.title"
    ),
  };
  const selectFileResult = await migrationHandlerDeps.selectFile(selectFileConfig);
  if (selectFileResult.isErr() || selectFileResult.value.type !== "success") {
    migrationHandlerDeps.sendTelemetryErrorEvent(
      TelemetryEvent.MigrateTeamsManifest,
      userCancelError
    );
    return ok(null);
  }
  const manifestPath = selectFileResult.value.result as string;

  const progressBar = migrationHandlerDeps.createProgressBar(
    migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.progressTitle"),
    1
  );
  await progressBar.start();

  const migrationHandler = migrationHandlerDeps.createMigrationHandler(manifestPath);
  let result: Result<null, FxError> = ok(null);

  try {
    // Update Teams manifest
    await progressBar.next(
      migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.updateManifest")
    );
    VsCodeLogInstance.info(
      migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.updateManifest")
    );
    result = await migrationHandler.updateManifest();
    if (result.isErr()) {
      throw result.error;
    }
  } catch (error) {
    result = migrationHandlerDeps.wrapError(error as Error);
  }

  if (result.isErr()) {
    await progressBar.end(false);
    void migrationHandlerDeps.showError(result.error);
    migrationHandlerDeps.sendTelemetryErrorEvent(TelemetryEvent.MigrateTeamsManifest, result.error);
  } else {
    await progressBar.end(true);
    void migrationHandlerDeps.showMessage(
      "info",
      util.format(
        migrationHandlerDeps.localize("teamstoolkit.migrateTeamsManifest.success"),
        manifestPath
      ),
      false
    );
    migrationHandlerDeps.sendTelemetryEvent(TelemetryEvent.MigrateTeamsManifest, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
  }
  return result;
}
