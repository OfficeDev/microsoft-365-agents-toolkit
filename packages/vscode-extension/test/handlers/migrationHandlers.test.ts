import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { ProgressHandler } from "@microsoft/vscode-ui";
import { assert } from "chai";
import VsCodeLogInstance from "../../src/commonlib/log";
import * as errorCommon from "../../src/error/common";
import { vi } from "vitest";
import {
  migrateTeamsManifestHandler,
  migrateTeamsTabAppHandler,
} from "../../src/handlers/migrationHandler";
import { migrationHandlerDeps } from "../../src/handlers/migrationHandler";
import { TeamsAppMigrationHandler } from "../../src/migration/migrationHandler";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as localizeUtils from "../../src/utils/localizeUtils";

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
});
