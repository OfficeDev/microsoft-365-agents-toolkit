/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import axios from "axios";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { ok, err } from "neverthrow";
import * as retryModule from "../../../../src/http/retry";
import { createMockContext } from "../../testHelper";
import { DriverRegistry } from "../../../../src/drivers/registry";
import { builtinDrivers } from "../../../../src/drivers/builtin";
import { oauthRegisterDriver } from "../../../../src/drivers/builtin/oauth/register";
import { apiKeyRegisterDriver } from "../../../../src/drivers/builtin/apiKey/register";

const FAKE_TOKEN = "fake-m365-token";

describe("Auth plugin drivers (oauth + apiKey)", () => {
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

  // ── Registration ───────────────────────────────────────────────

  describe("registration", () => {
    it("all 22 builtin drivers are registered", () => {
      expect(builtinDrivers).to.have.lengthOf(22);
    });

    it("registry contains oauth/register and apiKey/register", () => {
      const registry = new DriverRegistry();
      for (const d of builtinDrivers) {
        registry.register(d);
      }
      expect(registry.get("oauth/register")).to.not.be.undefined;
      expect(registry.get("apiKey/register")).to.not.be.undefined;
    });
  });

  // ── oauth/register driver ─────────────────────────────────────

  describe("oauth/register driver", () => {
    it("has correct metadata", () => {
      expect(oauthRegisterDriver.id).to.equal("oauth/register");
      expect(oauthRegisterDriver.name).to.equal("Register OAuth Configuration");
    });

    it("validates name is required", () => {
      const result = oauthRegisterDriver.validateFn!({
        name: "",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates name max length", () => {
      const result = oauthRegisterDriver.validateFn!({
        name: "x".repeat(129),
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates flow must be authorizationCode", () => {
      const result = oauthRegisterDriver.validateFn!({
        name: "test-oauth",
        appId: "app-123",
        flow: "implicit",
        clientId: "client-123",
        baseUrl: "https://example.com",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates baseUrl must be HTTPS", () => {
      const result = oauthRegisterDriver.validateFn!({
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "http://example.com",
      });
      expect(result.isErr()).to.be.true;
    });

    it("accepts valid config", () => {
      const result = oauthRegisterDriver.validateFn!({
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
      });
      expect(result.isOk()).to.be.true;
    });

    it("skips creation when existing config ID is found", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      // Mock getOauthRegistration returning existing config
      mockAxios.get.resolves({
        status: 200,
        data: { oAuthConfigId: "existing-config-id" },
      });

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
        existingConfigurationId: "existing-config-id",
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs.configurationId).to.equal("existing-config-id");
      }
      // Should have called GET, not POST
      expect(mockAxios.post.called).to.be.false;
    });

    it("returns error when Custom provider missing authorizationUrl", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
        identityProvider: "Custom",
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("MissingAuthorizationUrl");
      }
    });

    it("returns error when Custom provider missing tokenUrl", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
        identityProvider: "Custom",
        authorizationUrl: "https://example.com/authorize",
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("MissingTokenUrl");
      }
    });

    it("creates OAuth config for Custom provider", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      mockAxios.post.resolves({
        status: 200,
        data: {
          configurationRegistrationId: { oAuthConfigId: "new-config-123" },
          resourceIdentifierUri: "api://example.com",
        },
      });

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        clientSecret: "my-secret",
        baseUrl: "https://example.com",
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scope: "read,write",
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs.configurationId).to.equal("new-config-123");
      }
      expect(mockAxios.post.calledOnce).to.be.true;

      // Verify the registration payload
      const payload = mockAxios.post.firstCall.args[1];
      expect(payload.description).to.equal("test-oauth");
      expect(payload.clientId).to.equal("client-123");
      expect(payload.clientSecret).to.equal("my-secret");
      expect(payload.identityProvider).to.equal("Custom");
      expect(payload.authorizationEndpoint).to.equal("https://example.com/authorize");
      expect(payload.tokenExchangeEndpoint).to.equal("https://example.com/token");
      expect(payload.scopes).to.deep.equal(["read", "write"]);
      expect(payload.targetUrlsShouldStartWith).to.deep.equal(["https://example.com"]);
    });

    it("creates OAuth config for MicrosoftEntra provider", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      mockAxios.post.resolves({
        status: 200,
        data: {
          configurationRegistrationId: { oAuthConfigId: "entra-config-456" },
          resourceIdentifierUri: "api://entra.example.com",
        },
      });

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-entra-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
        identityProvider: "MicrosoftEntra",
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs.configurationId).to.equal("entra-config-456");
      }

      const payload = mockAxios.post.firstCall.args[1];
      expect(payload.identityProvider).to.equal("MicrosoftEntra");
      expect(payload.clientSecret).to.equal("");
    });

    it("accepts config without baseUrl when apiSpecPath is provided", () => {
      const result = oauthRegisterDriver.validateFn!({
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        apiSpecPath: "./appPackage/apiSpecificationFile/repair.yml",
        identityProvider: "MicrosoftEntra",
      });
      expect(result.isOk()).to.be.true;
    });

    it("derives domain from apiSpecPath when baseUrl is absent", async () => {
      const ctx = createMockContext();
      const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "oauth-test-"));
      ctx.projectPath = tmpDir;
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      // Write a real YAML spec file with a server URL
      const specDir = path.join(tmpDir, "appPackage", "apiSpecificationFile");
      await fs.mkdir(specDir, { recursive: true });
      const specContent = `
openapi: 3.0.0
info:
  title: Test
  version: 1.0.0
servers:
  - url: https://myserver.azurewebsites.net/api
paths: {}
`;
      await fs.writeFile(path.join(specDir, "repair.yml"), specContent, "utf-8");

      mockAxios.post.resolves({
        status: 200,
        data: {
          configurationRegistrationId: { oAuthConfigId: "spec-derived-config" },
          resourceIdentifierUri: "api://myserver.azurewebsites.net",
        },
      });

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        apiSpecPath: "./appPackage/apiSpecificationFile/repair.yml",
        identityProvider: "MicrosoftEntra",
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs.configurationId).to.equal("spec-derived-config");
      }

      const payload = mockAxios.post.firstCall.args[1];
      expect(payload.targetUrlsShouldStartWith).to.deep.equal([
        "https://myserver.azurewebsites.net",
      ]);
    });

    it("returns MissingBaseUrl when neither baseUrl nor apiSpecPath provided", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        identityProvider: "MicrosoftEntra",
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("MissingBaseUrl");
      }
    });

    it("returns error when token acquisition fails", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox
        .stub()
        .resolves(err({ message: "not signed in" }));

      const result = await oauthRegisterDriver.executeFn(ctx, {
        name: "test-oauth",
        appId: "app-123",
        flow: "authorizationCode",
        clientId: "client-123",
        baseUrl: "https://example.com",
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("TokenAcquisitionError");
      }
    });
  });

  // ── apiKey/register driver ────────────────────────────────────

  describe("apiKey/register driver", () => {
    it("has correct metadata", () => {
      expect(apiKeyRegisterDriver.id).to.equal("apiKey/register");
      expect(apiKeyRegisterDriver.name).to.equal("Register API Key");
    });

    it("validates name is required", () => {
      const result = apiKeyRegisterDriver.validateFn!({
        name: "",
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "secret12345",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates name max length", () => {
      const result = apiKeyRegisterDriver.validateFn!({
        name: "x".repeat(129),
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "secret12345",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates baseUrl must be HTTPS", () => {
      const result = apiKeyRegisterDriver.validateFn!({
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "http://example.com",
        primaryClientSecret: "secret12345",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates secret min length", () => {
      const result = apiKeyRegisterDriver.validateFn!({
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "short",
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates secret max length", () => {
      const result = apiKeyRegisterDriver.validateFn!({
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "x".repeat(513),
      });
      expect(result.isErr()).to.be.true;
    });

    it("accepts valid config", () => {
      const result = apiKeyRegisterDriver.validateFn!({
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "my-secret-12345",
      });
      expect(result.isOk()).to.be.true;
    });

    it("skips creation when existing registration ID is found", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));

      mockAxios.get.resolves({
        status: 200,
        data: { id: "existing-reg-id" },
      });

      const result = await apiKeyRegisterDriver.executeFn(ctx, {
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "my-secret-12345",
        existingRegistrationId: "existing-reg-id",
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs.registrationId).to.equal("existing-reg-id");
      }
      expect(mockAxios.post.called).to.be.false;
    });

    it("returns error when no client secret provided", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));
      (ctx.auth.m365TokenProvider as any).getJsonObject = sandbox
        .stub()
        .resolves(ok({ oid: "user-oid-123" }));

      const result = await apiKeyRegisterDriver.executeFn(ctx, {
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://example.com",
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("NoClientSecretProvided");
      }
    });

    it("creates API key registration with primary secret", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));
      (ctx.auth.m365TokenProvider as any).getJsonObject = sandbox
        .stub()
        .resolves(ok({ oid: "user-oid-123" }));

      mockAxios.post.resolves({
        status: 200,
        data: { id: "new-reg-456" },
      });

      const result = await apiKeyRegisterDriver.executeFn(ctx, {
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://api.example.com",
        primaryClientSecret: "my-primary-secret",
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs.registrationId).to.equal("new-reg-456");
      }
      expect(mockAxios.post.calledOnce).to.be.true;

      const payload = mockAxios.post.firstCall.args[1];
      expect(payload.description).to.equal("test-apikey");
      expect(payload.targetUrlsShouldStartWith).to.deep.equal(["https://api.example.com"]);
      expect(payload.clientSecrets).to.have.lengthOf(1);
      expect(payload.clientSecrets[0].value).to.equal("my-primary-secret");
      expect(payload.clientSecrets[0].priority).to.equal(0);
      expect(payload.manageableByUsers).to.have.lengthOf(1);
      expect(payload.manageableByUsers[0].userId).to.equal("user-oid-123");
    });

    it("creates API key with both primary and secondary secrets", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));
      (ctx.auth.m365TokenProvider as any).getJsonObject = sandbox
        .stub()
        .resolves(ok({ oid: "user-oid-123" }));

      mockAxios.post.resolves({
        status: 200,
        data: { id: "dual-reg-789" },
      });

      const result = await apiKeyRegisterDriver.executeFn(ctx, {
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://api.example.com",
        primaryClientSecret: "primary-secret-abc",
        secondaryClientSecret: "secondary-secret-xyz",
      });

      expect(result.isOk()).to.be.true;
      const payload = mockAxios.post.firstCall.args[1];
      expect(payload.clientSecrets).to.have.lengthOf(2);
      expect(payload.clientSecrets[0].priority).to.equal(0);
      expect(payload.clientSecrets[1].priority).to.equal(1);
    });

    it("returns error when token acquisition fails", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox
        .stub()
        .resolves(err({ message: "not signed in" }));

      const result = await apiKeyRegisterDriver.executeFn(ctx, {
        name: "test-apikey",
        appId: "app-123",
        baseUrl: "https://example.com",
        primaryClientSecret: "my-secret-12345",
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("TokenAcquisitionError");
      }
    });

    it("handles SpecificApp applicableToApps", async () => {
      const ctx = createMockContext();
      (ctx.auth.m365TokenProvider as any).getAccessToken = sandbox.stub().resolves(ok(FAKE_TOKEN));
      (ctx.auth.m365TokenProvider as any).getJsonObject = sandbox
        .stub()
        .resolves(ok({ oid: "user-oid-123" }));

      mockAxios.post.resolves({
        status: 200,
        data: { id: "specific-app-reg" },
      });

      const result = await apiKeyRegisterDriver.executeFn(ctx, {
        name: "test-apikey",
        appId: "my-specific-app",
        baseUrl: "https://api.example.com",
        primaryClientSecret: "my-secret-12345",
        applicableToApps: "SpecificApp",
      });

      expect(result.isOk()).to.be.true;
      const payload = mockAxios.post.firstCall.args[1];
      expect(payload.applicableToApps).to.equal("SpecificApp");
      expect(payload.specificAppId).to.equal("my-specific-app");
    });
  });
});
