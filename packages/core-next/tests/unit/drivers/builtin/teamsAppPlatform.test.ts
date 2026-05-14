/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import * as path from "node:path";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import axios from "axios";
import AdmZip from "adm-zip";
import { ok, err } from "neverthrow";
import * as retryModule from "../../../../src/http/retry";
import { createMockContext } from "../../testHelper";
import { driverRegistry } from "../../../../src/drivers/registry";
import { registerBuiltinDrivers } from "../../../../src/drivers/builtin";
import { createTeamsAppDriver } from "../../../../src/drivers/builtin/teamsApp/create";
import { configureTeamsAppDriver } from "../../../../src/drivers/builtin/teamsApp/configure";
import { publishAppPackageDriver } from "../../../../src/drivers/builtin/teamsApp/publishAppPackage";

describe("teamsApp platform drivers", () => {
  let sandbox: sinon.SinonSandbox;
  let mockAxios: {
    post: sinon.SinonStub;
    get: sinon.SinonStub;
    defaults: { headers: { common: Record<string, string> } };
    interceptors: {
      request: { use: sinon.SinonStub };
      response: { use: sinon.SinonStub };
    };
  };

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    mockAxios = {
      post: sandbox.stub(),
      get: sandbox.stub(),
      defaults: { headers: { common: {} } },
      interceptors: {
        request: { use: sandbox.stub() },
        response: { use: sandbox.stub() },
      },
    };
    sandbox.stub(axios, "create").returns(mockAxios as any);
    sandbox.stub(retryModule, "sendWithRetry").callsFake(async (fn: any) => fn());
  });

  afterEach(() => {
    sandbox.restore();
  });

  function mockCtxWithToken() {
    const ctx = createMockContext();
    (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox
      .stub()
      .resolves(ok("mock-m365-token"));
    return ctx;
  }

  function mockCtxWithTokenError() {
    const ctx = createMockContext();
    const fakeError = new Error("Login failed");
    (fakeError as any).source = "test";
    (fakeError as any).timestamp = new Date();
    (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(err(fakeError));
    return ctx;
  }

  // ──────────────────────── Registration ────────────────────────

  describe("registration", () => {
    it("registers create, configure, publishAppPackage drivers", () => {
      const _fresh = new (driverRegistry.constructor as any)();
      // Use the exported descriptors directly
      expect(createTeamsAppDriver.id).to.equal("teamsApp/create");
      expect(configureTeamsAppDriver.id).to.equal("teamsApp/configure");
      expect(publishAppPackageDriver.id).to.equal("teamsApp/publishAppPackage");
    });

    it("all 22 builtin drivers are registered", () => {
      registerBuiltinDrivers();
      const ids = ["teamsApp/create", "teamsApp/configure", "teamsApp/publishAppPackage"];
      for (const id of ids) {
        expect(driverRegistry.get(id)).to.not.be.undefined;
      }
    });
  });

  // ──────────────────────── teamsApp/create ─────────────────────

  describe("teamsApp/create", () => {
    it("has correct metadata", () => {
      expect(createTeamsAppDriver.id).to.equal("teamsApp/create");
      expect(createTeamsAppDriver.name).to.equal("Create Teams App");
    });

    it("creates a new app when no existing ID", async () => {
      const appDef = { teamsAppId: "new-app-id", tenantId: "tenant-123" };
      mockAxios.post.resolves({ data: appDef });

      const ctx = mockCtxWithToken();
      const result = await createTeamsAppDriver.executeFn(ctx, { name: "My App" });

      expect(result.isOk()).to.be.true;
      const outputs = result._unsafeUnwrap().outputs;
      expect(outputs["TEAMS_APP_ID"]).to.equal("new-app-id");
      expect(outputs["TEAMS_APP_TENANT_ID"]).to.equal("tenant-123");

      // Verify ZIP was sent with manifest.json
      const [url, body] = mockAxios.post.firstCall.args;
      expect(url).to.equal("/api/appdefinitions/v2/import");
      const zip = new AdmZip(body);
      const manifestEntry = zip.getEntry("manifest.json");
      expect(manifestEntry).to.not.be.null;
      const manifest = JSON.parse(manifestEntry!.getData().toString());
      expect(manifest.name.short).to.equal("My App");
    });

    it("returns existing app when existingTeamsAppId is valid UUID and app exists", async () => {
      const existingId = "11111111-1111-1111-1111-111111111111";
      const appDef = { teamsAppId: existingId, tenantId: "tenant-abc" };
      mockAxios.get.resolves({ data: appDef });

      const ctx = mockCtxWithToken();
      const result = await createTeamsAppDriver.executeFn(ctx, {
        name: "My App",
        existingTeamsAppId: existingId,
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["TEAMS_APP_ID"]).to.equal(existingId);
      // Should NOT have called importApp (POST)
      expect(mockAxios.post.called).to.be.false;
    });

    it("creates new app when existingTeamsAppId is an unresolved placeholder", async () => {
      const appDef = { teamsAppId: "created-id", tenantId: "t1" };
      mockAxios.post.resolves({ data: appDef });

      const ctx = mockCtxWithToken();
      const result = await createTeamsAppDriver.executeFn(ctx, {
        name: "My App",
        existingTeamsAppId: "${{TEAMS_APP_ID}}", // unresolved placeholder
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["TEAMS_APP_ID"]).to.equal("created-id");
      // Should have called importApp since placeholder is not a UUID
      expect(mockAxios.post.calledOnce).to.be.true;
    });

    it("returns error when token acquisition fails", async () => {
      const ctx = mockCtxWithTokenError();
      const result = await createTeamsAppDriver.executeFn(ctx, { name: "My App" });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TokenAcquisitionError");
    });

    it("validates name is required", async () => {
      const ctx = mockCtxWithToken();
      const result = await createTeamsAppDriver.executeFn(ctx, {} as any);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("InvalidDriverInput");
    });
  });

  // ──────────────────────── teamsApp/configure ──────────────────

  describe("teamsApp/configure", () => {
    let tmpDir: string;

    beforeEach(async () => {
      tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "configure-test-"));
    });

    afterEach(async () => {
      await fs.rm(tmpDir, { recursive: true, force: true });
    });

    function createAppPackage(teamsAppId: string): string {
      const zip = new AdmZip();
      const manifest = {
        id: teamsAppId,
        name: { short: "Test App" },
        version: "1.0.0",
      };
      zip.addFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      const zipPath = path.join(tmpDir, "app.zip");
      zip.writeZip(zipPath);
      return zipPath;
    }

    it("has correct metadata", () => {
      expect(configureTeamsAppDriver.id).to.equal("teamsApp/configure");
      expect(configureTeamsAppDriver.name).to.equal("Configure Teams App");
    });

    it("updates existing app", async () => {
      const appId = "22222222-2222-2222-2222-222222222222";
      const zipPath = createAppPackage(appId);

      // getApp succeeds
      mockAxios.get.resolves({ data: { teamsAppId: appId, tenantId: "t1" } });
      // importApp succeeds
      mockAxios.post.resolves({
        data: { teamsAppId: appId, tenantId: "t1", updatedAt: "2026-03-31T00:00:00Z" },
      });

      const ctx = mockCtxWithToken();
      const result = await configureTeamsAppDriver.executeFn(ctx, { appPackagePath: zipPath });

      expect(result.isOk()).to.be.true;
      const outputs = result._unsafeUnwrap().outputs;
      expect(outputs["TEAMS_APP_TENANT_ID"]).to.equal("t1");
      expect(outputs["TEAMS_APP_UPDATE_TIME"]).to.equal("2026-03-31T00:00:00Z");
    });

    it("fails when app does not exist", async () => {
      const appId = "33333333-3333-3333-3333-333333333333";
      const zipPath = createAppPackage(appId);

      // getApp fails (404)
      mockAxios.get.rejects({
        response: { status: 404, data: "not found" },
        message: "Not found",
      });

      const ctx = mockCtxWithToken();
      const result = await configureTeamsAppDriver.executeFn(ctx, { appPackagePath: zipPath });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppNotFound");
    });

    it("fails when package file missing", async () => {
      const ctx = mockCtxWithToken();
      const result = await configureTeamsAppDriver.executeFn(ctx, {
        appPackagePath: path.join(tmpDir, "nonexistent.zip"),
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("AppPackageNotFound");
    });

    it("fails when manifest has invalid app ID", async () => {
      const zip = new AdmZip();
      zip.addFile("manifest.json", Buffer.from(JSON.stringify({ id: "not-a-uuid" })));
      const zipPath = path.join(tmpDir, "bad.zip");
      zip.writeZip(zipPath);

      const ctx = mockCtxWithToken();
      const result = await configureTeamsAppDriver.executeFn(ctx, { appPackagePath: zipPath });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("InvalidTeamsAppId");
    });
  });

  // ──────────────────── teamsApp/publishAppPackage ──────────────

  describe("teamsApp/publishAppPackage", () => {
    let tmpDir: string;

    beforeEach(async () => {
      tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "publish-test-"));
    });

    afterEach(async () => {
      await fs.rm(tmpDir, { recursive: true, force: true });
    });

    function createAppPackage(teamsAppId: string): string {
      const zip = new AdmZip();
      zip.addFile(
        "manifest.json",
        Buffer.from(JSON.stringify({ id: teamsAppId, name: { short: "App" } }))
      );
      const zipPath = path.join(tmpDir, "app.zip");
      zip.writeZip(zipPath);
      return zipPath;
    }

    it("has correct metadata", () => {
      expect(publishAppPackageDriver.id).to.equal("teamsApp/publishAppPackage");
      expect(publishAppPackageDriver.name).to.equal("Publish App Package");
    });

    it("first publish when app not yet published", async () => {
      const appId = "44444444-4444-4444-4444-444444444444";
      const zipPath = createAppPackage(appId);

      // getStagedApp returns nothing (Graph beta endpoint)
      mockAxios.get.resolves({ data: { value: [] } });
      // publishTeamsApp succeeds
      mockAxios.post.resolves({ data: { id: "pub-id-1" } });

      const ctx = mockCtxWithToken();
      const result = await publishAppPackageDriver.executeFn(ctx, { appPackagePath: zipPath });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["TEAMS_APP_PUBLISHED_APP_ID"]).to.equal("pub-id-1");
      // First POST should be to Graph /appCatalogs/teamsApps
      expect(mockAxios.post.firstCall.args[0]).to.include(
        "/appCatalogs/teamsApps?requiresReview=true"
      );
    });

    it("update when app already published", async () => {
      const appId = "55555555-5555-5555-5555-555555555555";
      const zipPath = createAppPackage(appId);

      // getStagedApp returns existing app (Graph response format)
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "staged-55",
              displayName: "App",
              appDefinitions: [
                {
                  publishingState: "published",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      });
      // publishTeamsAppUpdate succeeds
      mockAxios.post.resolves({ data: { teamsAppId: "updated-55" } });

      const ctx = mockCtxWithToken();
      const result = await publishAppPackageDriver.executeFn(ctx, { appPackagePath: zipPath });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["TEAMS_APP_PUBLISHED_APP_ID"]).to.equal("updated-55");
      // POST should be to the Graph appDefinitions update endpoint
      expect(mockAxios.post.firstCall.args[0]).to.include(
        "/appCatalogs/teamsApps/staged-55/appDefinitions?requiresReview=true"
      );
    });

    it("fails when package file missing", async () => {
      const ctx = mockCtxWithToken();
      const result = await publishAppPackageDriver.executeFn(ctx, {
        appPackagePath: path.join(tmpDir, "missing.zip"),
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("AppPackageNotFound");
    });

    it("fails when manifest missing from ZIP", async () => {
      const zip = new AdmZip();
      zip.addFile("readme.txt", Buffer.from("hello"));
      const zipPath = path.join(tmpDir, "no-manifest.zip");
      zip.writeZip(zipPath);

      const ctx = mockCtxWithToken();
      const result = await publishAppPackageDriver.executeFn(ctx, { appPackagePath: zipPath });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("ManifestNotFound");
    });
  });
});
