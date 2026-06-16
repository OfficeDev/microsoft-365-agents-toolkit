import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { ProgressHandler } from "@microsoft/vscode-ui";
import { assert } from "chai";
import * as vscode from "vscode";
import VsCodeLogInstance from "../../src/commonlib/log";
import * as errorCommon from "../../src/error/common";
import { vi } from "vitest";
import {
  migrateTeamsManifestHandler,
  migrateTeamsTabAppHandler,
} from "../../src/handlers/migrationHandler";
import { migrationHandlerOps } from "../../src/handlers/migrationHandler";
import { TeamsAppMigrationHandler } from "../../src/migration/migrationHandler";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as localizeUtils from "../../src/utils/localizeUtils";
import { mockValue } from "../mocks/vitestMockUtils";

const migrationHandlerDeps = migrationHandlerOps;

describe("Migration handlers", () => {
  beforeEach(() => {
    vi.spyOn(migrationHandlerDeps, "sendTelemetryEvent");
  });

  describe("migrateTeamsTabAppHandler", () => {
    it("happy path", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsTabApp.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFolder").mockResolvedValue(
        ok({ type: "success", result: "test" })
      );
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updatePackageJson").mockResolvedValue(ok(true));
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updateCodes").mockResolvedValue(ok([]));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
    });

    it("happy path: failed files", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsTabApp.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFolder").mockResolvedValue(
        ok({ type: "success", result: "test" })
      );
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      const warningStub = vi.spyOn(VsCodeLogInstance, "warning");
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updatePackageJson").mockResolvedValue(ok(true));
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updateCodes").mockResolvedValue(
        ok(["test1", "test2"])
      );

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(warningStub.calledOnce);
    });

    it("error", async () => {
      const sendTelemetryErrorEventStub = vi.spyOn(migrationHandlerDeps, "sendTelemetryErrorEvent");
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsTabApp.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFolder").mockResolvedValue(
        ok({ type: "success", result: "test" })
      );
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updatePackageJson").mockResolvedValue(ok(true));
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updateCodes").mockResolvedValue(
        err({ foo: "bar" } as any)
      );

      const result = await migrateTeamsTabAppHandler();

      assert.isTrue(result.isErr());
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("user cancel", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const sendTelemetryErrorEventStub = vi.spyOn(migrationHandlerDeps, "sendTelemetryErrorEvent");
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsTabApp.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFolder").mockResolvedValue(ok({ type: "skip" }));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("user cancel: skip folder selection", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const sendTelemetryErrorEventStub = vi.spyOn(migrationHandlerDeps, "sendTelemetryErrorEvent");
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(ok("cancel"));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("no change in package.json", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsTabApp.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFolder").mockResolvedValue(
        ok({ type: "success", result: "test" })
      );
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      vi.spyOn(VsCodeLogInstance, "warning").mockReturnValue();
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updatePackageJson").mockResolvedValue(
        ok(false)
      );

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
    });
  });

  describe("migrateTeamsManifestHandler", () => {
    it("happy path", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsManifest.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFile").mockResolvedValue(
        ok({ type: "success", result: "test" })
      );
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updateManifest").mockResolvedValue(ok(null));

      const result = await migrateTeamsManifestHandler();

      assert.deepEqual(result, ok(null));
    });

    it("user cancel: skip file selection", async () => {
      const sendTelemetryErrorEventStub = vi.spyOn(migrationHandlerDeps, "sendTelemetryErrorEvent");
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsManifest.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFile").mockResolvedValue(ok({ type: "skip" }));
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updateManifest").mockResolvedValue(ok(null));

      const result = await migrateTeamsManifestHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("error", async () => {
      vi.spyOn(migrationHandlerDeps, "localize").mockImplementation((key: string) => key);
      const sendTelemetryErrorEventStub = vi.spyOn(migrationHandlerDeps, "sendTelemetryErrorEvent");
      const progressHandler = new ProgressHandler("title", 1);
      vi.spyOn(migrationHandlerDeps, "showMessage").mockResolvedValue(
        ok("teamstoolkit.migrateTeamsManifest.upgrade")
      );
      vi.spyOn(migrationHandlerDeps, "selectFile").mockResolvedValue(
        ok({ type: "success", result: "test" })
      );
      vi.spyOn(migrationHandlerDeps, "createProgressBar").mockReturnValue(progressHandler);
      vi.spyOn(VsCodeLogInstance, "info").mockReturnValue();
      vi.spyOn(TeamsAppMigrationHandler.prototype, "updateManifest").mockResolvedValue(
        err(new UserError("source", "name", ""))
      );
      vi.spyOn(migrationHandlerDeps, "showError").mockImplementation(async () => {});

      const result = await migrateTeamsManifestHandler();

      assert.isTrue(result.isErr());
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });
  });

  describe("migrationHandlerDeps delegation", () => {
    beforeEach(() => {
      mockValue(vsc_ui, "VS_CODE_UI", new vsc_ui.VsCodeUI(vscode.window, vscode.workspace));
    });

    it("delegates showMessage/selectFolder/selectFile/createProgressBar", async () => {
      vi.spyOn(vsc_ui.VS_CODE_UI, "showMessage").mockResolvedValue(ok("OK"));
      vi.spyOn(vsc_ui.VS_CODE_UI, "selectFolder").mockResolvedValue(ok({ type: "skip" }));
      vi.spyOn(vsc_ui.VS_CODE_UI, "selectFile").mockResolvedValue(ok({ type: "skip" }));
      vi.spyOn(vsc_ui.VS_CODE_UI, "createProgressBar").mockReturnValue(new ProgressHandler("t", 1));

      const msg = await migrationHandlerDeps.showMessage("info", "m", false, "OK");
      const folder = await migrationHandlerDeps.selectFolder({ name: "n", title: "t" });
      const file = await migrationHandlerDeps.selectFile({ name: "n", title: "t" });
      const progress = migrationHandlerDeps.createProgressBar("title", 1);

      assert.isTrue(msg.isOk());
      assert.isTrue(folder.isOk());
      assert.isTrue(file.isOk());
      assert.exists(progress);
    });

    it("delegates wrapError/showError/localize/createMigrationHandler", async () => {
      vi.spyOn(errorCommon, "wrapError").mockReturnValue(err({ foo: "bar" } as any));
      vi.spyOn(errorCommon, "showError").mockResolvedValue();
      vi.spyOn(localizeUtils, "localize").mockReturnValue("localized");

      const wrapped = migrationHandlerDeps.wrapError(new Error("x"));
      await migrationHandlerDeps.showError({} as any);
      const localized = migrationHandlerDeps.localize("key");
      const handler = migrationHandlerDeps.createMigrationHandler("C:\\test");

      assert.isTrue(wrapped.isErr());
      assert.equal(localized, "localized");
      assert.instanceOf(handler, TeamsAppMigrationHandler);
    });
  });
});
