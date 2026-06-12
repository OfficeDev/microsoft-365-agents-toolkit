import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { ProgressHandler } from "@microsoft/vscode-ui";
import * as sinon from "sinon";
import { assert } from "chai";
import VsCodeLogInstance from "../../src/commonlib/log";
import * as errorCommon from "../../src/error/common";
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
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(migrationHandlerDeps, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("migrateTeamsTabAppHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsTabApp.upgrade"));
      sandbox
        .stub(migrationHandlerDeps, "selectFolder")
        .resolves(ok({ type: "success", result: "test" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updateCodes").resolves(ok([]));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
    });

    it("happy path: failed files", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsTabApp.upgrade"));
      sandbox
        .stub(migrationHandlerDeps, "selectFolder")
        .resolves(ok({ type: "success", result: "test" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      const warningStub = sandbox.stub(VsCodeLogInstance, "warning");
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sandbox
        .stub(TeamsAppMigrationHandler.prototype, "updateCodes")
        .resolves(ok(["test1", "test2"]));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(warningStub.calledOnce);
    });

    it("error", async () => {
      const sendTelemetryErrorEventStub = sandbox.stub(
        migrationHandlerDeps,
        "sendTelemetryErrorEvent"
      );
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsTabApp.upgrade"));
      sandbox
        .stub(migrationHandlerDeps, "selectFolder")
        .resolves(ok({ type: "success", result: "test" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(true));
      sandbox
        .stub(TeamsAppMigrationHandler.prototype, "updateCodes")
        .resolves(err({ foo: "bar" } as any));

      const result = await migrateTeamsTabAppHandler();

      assert.isTrue(result.isErr());
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("user cancel", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sandbox.stub(
        migrationHandlerDeps,
        "sendTelemetryErrorEvent"
      );
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsTabApp.upgrade"));
      sandbox.stub(migrationHandlerDeps, "selectFolder").resolves(ok({ type: "skip" }));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("user cancel: skip folder selection", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sandbox.stub(
        migrationHandlerDeps,
        "sendTelemetryErrorEvent"
      );
      sandbox.stub(migrationHandlerDeps, "showMessage").resolves(ok("cancel"));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("no change in package.json", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsTabApp.upgrade"));
      sandbox
        .stub(migrationHandlerDeps, "selectFolder")
        .resolves(ok({ type: "success", result: "test" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(VsCodeLogInstance, "warning").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updatePackageJson").resolves(ok(false));

      const result = await migrateTeamsTabAppHandler();

      assert.deepEqual(result, ok(null));
    });
  });

  describe("migrateTeamsManifestHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsManifest.upgrade"));
      sandbox
        .stub(migrationHandlerDeps, "selectFile")
        .resolves(ok({ type: "success", result: "test" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updateManifest").resolves(ok(null));

      const result = await migrateTeamsManifestHandler();

      assert.deepEqual(result, ok(null));
    });

    it("user cancel: skip file selection", async () => {
      const sendTelemetryErrorEventStub = sandbox.stub(
        migrationHandlerDeps,
        "sendTelemetryErrorEvent"
      );
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsManifest.upgrade"));
      sandbox.stub(migrationHandlerDeps, "selectFile").resolves(ok({ type: "skip" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox.stub(TeamsAppMigrationHandler.prototype, "updateManifest").resolves(ok(null));

      const result = await migrateTeamsManifestHandler();

      assert.deepEqual(result, ok(null));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });

    it("error", async () => {
      sandbox.stub(migrationHandlerDeps, "localize").callsFake((key: string) => key);
      const sendTelemetryErrorEventStub = sandbox.stub(
        migrationHandlerDeps,
        "sendTelemetryErrorEvent"
      );
      const progressHandler = new ProgressHandler("title", 1);
      sandbox
        .stub(migrationHandlerDeps, "showMessage")
        .resolves(ok("teamstoolkit.migrateTeamsManifest.upgrade"));
      sandbox
        .stub(migrationHandlerDeps, "selectFile")
        .resolves(ok({ type: "success", result: "test" }));
      sandbox.stub(migrationHandlerDeps, "createProgressBar").returns(progressHandler);
      sandbox.stub(VsCodeLogInstance, "info").returns();
      sandbox
        .stub(TeamsAppMigrationHandler.prototype, "updateManifest")
        .resolves(err(new UserError("source", "name", "")));
      sandbox.stub(migrationHandlerDeps, "showError").callsFake(async () => {});

      const result = await migrateTeamsManifestHandler();

      assert.isTrue(result.isErr());
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });
  });
});
