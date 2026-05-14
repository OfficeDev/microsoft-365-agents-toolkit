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
import { ok, err } from "neverthrow";
import * as retryModule from "../../../../src/http/retry";
import { createMockContext } from "../../testHelper";
import { DriverRegistry } from "../../../../src/drivers/registry";
import { builtinDrivers } from "../../../../src/drivers/builtin";
import { createAadAppDriver } from "../../../../src/drivers/builtin/aadApp/create";
import { updateAadAppDriver } from "../../../../src/drivers/builtin/aadApp/update";
import { createBotAadAppDriver } from "../../../../src/drivers/builtin/botAadApp/create";
import { createBotFrameworkDriver } from "../../../../src/drivers/builtin/botFramework/create";

describe("Entra/Bot drivers", () => {
  let sandbox: sinon.SinonSandbox;
  let mockAxios: {
    post: sinon.SinonStub;
    get: sinon.SinonStub;
    patch: sinon.SinonStub;
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
      patch: sandbox.stub(),
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

  // Build a JWT with a tid claim for Graph token
  function fakeGraphToken(tid = "tenant-abc"): string {
    const header = Buffer.from(JSON.stringify({ alg: "none" })).toString("base64url");
    const payload = Buffer.from(JSON.stringify({ tid })).toString("base64url");
    return `${header}.${payload}.sig`;
  }

  function mockCtxWithGraphToken(tid = "tenant-abc") {
    const ctx = createMockContext();
    (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox
      .stub()
      .resolves(ok(fakeGraphToken(tid)));
    return ctx;
  }

  function mockCtxWithTokenError() {
    const ctx = createMockContext();
    const fakeError = new Error("token failed");
    (fakeError as any).source = "test";
    (fakeError as any).timestamp = new Date();
    (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(err(fakeError));
    return ctx;
  }

  // ──────────────────────── Registration ────────────────────────

  describe("registration", () => {
    it("registers all 4 new drivers", () => {
      const registry = new DriverRegistry();
      for (const driver of builtinDrivers) {
        registry.register(driver);
      }
      expect(registry.has("aadApp/create")).to.be.true;
      expect(registry.has("aadApp/update")).to.be.true;
      expect(registry.has("botAadApp/create")).to.be.true;
      expect(registry.has("botFramework/create")).to.be.true;
    });

    it("all 22 builtin drivers are registered", () => {
      expect(builtinDrivers).to.have.lengthOf(22);
    });
  });

  // ──────────────────────── aadApp/create ────────────────────────

  describe("aadApp/create", () => {
    it("has correct metadata", () => {
      expect(createAadAppDriver.id).to.equal("aadApp/create");
      expect(createAadAppDriver.name).to.equal("Create AAD App");
    });

    it("creates a new AAD app", async () => {
      mockAxios.post.resolves({
        data: {
          id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
          appId: "11111111-2222-3333-4444-555555555555",
          displayName: "TestApp",
        },
        status: 200,
      });

      const ctx = mockCtxWithGraphToken();
      const result = await createAadAppDriver.executeFn(ctx, {
        name: "TestApp",
        generateClientSecret: false,
      });

      expect(result.isOk()).to.be.true;
      const outputs = result._unsafeUnwrap().outputs;
      expect(outputs["AAD_APP_CLIENT_ID"]).to.equal("11111111-2222-3333-4444-555555555555");
      expect(outputs["AAD_APP_OBJECT_ID"]).to.equal("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
      expect(outputs["AAD_APP_TENANT_ID"]).to.equal("tenant-abc");
      expect(outputs["AAD_APP_OAUTH_AUTHORITY_HOST"]).to.equal("https://login.microsoftonline.com");
      expect(outputs["AAD_APP_OAUTH_AUTHORITY"]).to.equal(
        "https://login.microsoftonline.com/tenant-abc"
      );
    });

    it("skips creation when existingClientId is a valid UUID", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createAadAppDriver.executeFn(ctx, {
        name: "TestApp",
        generateClientSecret: false,
        existingClientId: "11111111-2222-3333-4444-555555555555",
        existingObjectId: "66666666-7777-8888-9999-aaaaaaaaaaaa",
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["AAD_APP_CLIENT_ID"]).to.equal(
        "11111111-2222-3333-4444-555555555555"
      );
      // Should not call Graph API
      expect(mockAxios.post.called).to.be.false;
    });

    it("generates client secret when requested", async () => {
      // First call: createAadApp, second call: generateClientSecret
      mockAxios.post
        .onFirstCall()
        .resolves({
          data: { id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", appId: "client-1" },
          status: 200,
        })
        .onSecondCall()
        .resolves({ data: { secretText: "my-secret" }, status: 200 });

      const ctx = mockCtxWithGraphToken();
      const result = await createAadAppDriver.executeFn(ctx, {
        name: "TestApp",
        generateClientSecret: true,
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["SECRET_AAD_APP_CLIENT_SECRET"]).to.equal("my-secret");
    });

    it("skips secret generation when existingClientSecret is set", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createAadAppDriver.executeFn(ctx, {
        name: "TestApp",
        generateClientSecret: true,
        existingClientId: "11111111-2222-3333-4444-555555555555",
        existingObjectId: "66666666-7777-8888-9999-aaaaaaaaaaaa",
        existingClientSecret: "already-have-one",
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["SECRET_AAD_APP_CLIENT_SECRET"]).to.equal(
        "already-have-one"
      );
    });

    it("returns error when token acquisition fails", async () => {
      const ctx = mockCtxWithTokenError();
      const result = await createAadAppDriver.executeFn(ctx, {
        name: "TestApp",
        generateClientSecret: false,
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TokenAcquisitionError");
    });

    it("validates name is required", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createAadAppDriver.executeFn(ctx, {
        name: "",
        generateClientSecret: false,
      });
      expect(result.isErr()).to.be.true;
    });
  });

  // ──────────────────────── aadApp/update ────────────────────────

  describe("aadApp/update", () => {
    let tmpDir: string;

    beforeEach(async () => {
      tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "aad-update-"));
    });

    afterEach(async () => {
      await fs.rm(tmpDir, { recursive: true, force: true });
    });

    it("has correct metadata", () => {
      expect(updateAadAppDriver.id).to.equal("aadApp/update");
      expect(updateAadAppDriver.name).to.equal("Update AAD App");
    });

    it("reads manifest and PATCHes the app", async () => {
      const manifest = {
        id: "obj-123",
        appId: "client-123",
        displayName: "Updated App",
        web: { redirectUris: ["https://example.com/callback"] },
      };
      await fs.writeFile(path.join(tmpDir, "aad.manifest.json"), JSON.stringify(manifest));

      mockAxios.patch.resolves({ data: {}, status: 204 });

      const ctx = mockCtxWithGraphToken();
      ctx.projectPath = tmpDir;

      const result = await updateAadAppDriver.executeFn(ctx, {
        manifestPath: "aad.manifest.json",
        outputFilePath: "build/aad.manifest.json",
      });

      expect(result.isOk()).to.be.true;
      expect(mockAxios.patch.calledOnce).to.be.true;
      // Verify output file was written
      const outputContent = await fs.readFile(
        path.join(tmpDir, "build", "aad.manifest.json"),
        "utf-8"
      );
      expect(JSON.parse(outputContent).displayName).to.equal("Updated App");
    });

    it("does two-phase update for preAuthorizedApplications", async () => {
      const manifest = {
        id: "obj-123",
        api: {
          oauth2PermissionScopes: [{ id: "scope-1", value: "access_as_user" }],
          preAuthorizedApplications: [
            { appId: "pre-auth-app", delegatedPermissionIds: ["scope-1"] },
          ],
        },
      };
      await fs.writeFile(path.join(tmpDir, "aad.manifest.json"), JSON.stringify(manifest));

      mockAxios.patch.resolves({ data: {}, status: 204 });

      const ctx = mockCtxWithGraphToken();
      ctx.projectPath = tmpDir;

      const result = await updateAadAppDriver.executeFn(ctx, {
        manifestPath: "aad.manifest.json",
        outputFilePath: "build/aad.manifest.json",
      });

      expect(result.isOk()).to.be.true;
      // Should PATCH twice: phase 1 (empty preAuth) + phase 2 (with preAuth)
      expect(mockAxios.patch.calledTwice).to.be.true;
    });

    it("fails when manifest file is missing", async () => {
      const ctx = mockCtxWithGraphToken();
      ctx.projectPath = tmpDir;

      const result = await updateAadAppDriver.executeFn(ctx, {
        manifestPath: "nonexistent.json",
        outputFilePath: "build/out.json",
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("ManifestFileNotFound");
    });

    it("fails when manifest has no id field", async () => {
      await fs.writeFile(
        path.join(tmpDir, "aad.manifest.json"),
        JSON.stringify({ displayName: "No ID" })
      );

      const ctx = mockCtxWithGraphToken();
      ctx.projectPath = tmpDir;

      const result = await updateAadAppDriver.executeFn(ctx, {
        manifestPath: "aad.manifest.json",
        outputFilePath: "build/out.json",
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("MissingObjectId");
    });

    it("fails when projectPath is undefined", async () => {
      const ctx = mockCtxWithGraphToken();
      ctx.projectPath = undefined;

      const result = await updateAadAppDriver.executeFn(ctx, {
        manifestPath: "aad.manifest.json",
        outputFilePath: "build/out.json",
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("MissingProjectPath");
    });
  });

  // ──────────────────────── botAadApp/create ────────────────────────

  describe("botAadApp/create", () => {
    it("has correct metadata", () => {
      expect(createBotAadAppDriver.id).to.equal("botAadApp/create");
      expect(createBotAadAppDriver.name).to.equal("Create Bot AAD App");
    });

    it("creates a new bot AAD app", async () => {
      mockAxios.post
        .onFirstCall()
        .resolves({
          data: {
            id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
            appId: "22222222-3333-4444-5555-666666666666",
          },
          status: 200,
        })
        .onSecondCall()
        .resolves({ data: { secretText: "bot-secret" }, status: 200 });

      const ctx = mockCtxWithGraphToken();
      const result = await createBotAadAppDriver.executeFn(ctx, {
        name: "MyBot",
      });

      expect(result.isOk()).to.be.true;
      const outputs = result._unsafeUnwrap().outputs;
      expect(outputs["BOT_ID"]).to.equal("22222222-3333-4444-5555-666666666666");
      expect(outputs["SECRET_BOT_PASSWORD"]).to.equal("bot-secret");
    });

    it("reuses existing when both botId and botPassword present", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createBotAadAppDriver.executeFn(ctx, {
        name: "MyBot",
        existingBotId: "11111111-2222-3333-4444-555555555555",
        existingBotPassword: "existing-pwd",
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs["BOT_ID"]).to.equal(
        "11111111-2222-3333-4444-555555555555"
      );
      expect(mockAxios.post.called).to.be.false;
    });

    it("errors when botId exists but password is empty", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createBotAadAppDriver.executeFn(ctx, {
        name: "MyBot",
        existingBotId: "11111111-2222-3333-4444-555555555555",
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("UnexpectedEmptyBotPassword");
    });

    it("uses AzureADMultipleOrgs audience", async () => {
      mockAxios.post
        .onFirstCall()
        .resolves({
          data: {
            id: "aaaaaaaa-0000-0000-0000-000000000000",
            appId: "bbbbbbbb-0000-0000-0000-000000000000",
          },
          status: 200,
        })
        .onSecondCall()
        .resolves({ data: { secretText: "s" }, status: 200 });

      const ctx = mockCtxWithGraphToken();
      await createBotAadAppDriver.executeFn(ctx, { name: "Bot" });

      const body = mockAxios.post.firstCall.args[1];
      expect(body.signInAudience).to.equal("AzureADMultipleOrgs");
    });

    it("returns error when token acquisition fails", async () => {
      const ctx = mockCtxWithTokenError();
      const result = await createBotAadAppDriver.executeFn(ctx, { name: "Bot" });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TokenAcquisitionError");
    });
  });

  // ──────────────────────── botFramework/create ────────────────────────

  describe("botFramework/create", () => {
    it("has correct metadata", () => {
      expect(createBotFrameworkDriver.id).to.equal("botFramework/create");
      expect(createBotFrameworkDriver.name).to.equal("Create Bot Framework Registration");
    });

    it("creates a new bot registration when none exists", async () => {
      // getBotRegistration returns 404
      mockAxios.get.rejects({ response: { status: 404 } });
      // createBotRegistration succeeds
      mockAxios.post.resolves({ data: {}, status: 200 });

      const ctx = mockCtxWithGraphToken(); // reuses same token mock, scopes differ but mock returns same
      const result = await createBotFrameworkDriver.executeFn(ctx, {
        botId: "11111111-2222-3333-4444-555555555555",
        name: "MyBot",
        messagingEndpoint: "https://example.com/api/messages",
      });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().outputs).to.deep.equal({});
      expect(mockAxios.post.calledOnce).to.be.true;
    });

    it("updates existing bot registration", async () => {
      // getBotRegistration returns existing
      mockAxios.get.resolves({
        data: {
          botId: "11111111-2222-3333-4444-555555555555",
          name: "OldBot",
          description: "old desc",
          iconUrl: "",
          messagingEndpoint: "https://old.com/api/messages",
          callingEndpoint: "",
          configuredChannels: ["msteams"],
        },
        status: 200,
      });
      // updateBotRegistration succeeds
      mockAxios.post.resolves({ data: {}, status: 200 });

      const ctx = mockCtxWithGraphToken();
      const result = await createBotFrameworkDriver.executeFn(ctx, {
        botId: "11111111-2222-3333-4444-555555555555",
        name: "UpdatedBot",
        messagingEndpoint: "https://new.com/api/messages",
      });

      expect(result.isOk()).to.be.true;
      // Should POST to update endpoint
      expect(mockAxios.post.calledOnce).to.be.true;
      const updateUrl = mockAxios.post.firstCall.args[0];
      expect(updateUrl).to.include("/api/botframework/11111111");
    });

    it("fails for invalid botId", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createBotFrameworkDriver.executeFn(ctx, {
        botId: "not-a-uuid",
        name: "Bot",
        messagingEndpoint: "https://example.com/api/messages",
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("InvalidBotId");
    });

    it("returns error when token acquisition fails", async () => {
      const ctx = mockCtxWithTokenError();
      const result = await createBotFrameworkDriver.executeFn(ctx, {
        botId: "11111111-2222-3333-4444-555555555555",
        name: "Bot",
        messagingEndpoint: "https://example.com/api/messages",
      });

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TokenAcquisitionError");
    });

    it("validates required fields", async () => {
      const ctx = mockCtxWithGraphToken();
      const result = await createBotFrameworkDriver.executeFn(ctx, {
        botId: "11111111-2222-3333-4444-555555555555",
        name: "",
        messagingEndpoint: "https://example.com/api/messages",
      });

      expect(result.isErr()).to.be.true;
    });
  });
});
