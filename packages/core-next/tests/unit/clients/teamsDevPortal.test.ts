/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import axios from "axios";
import * as retryModule from "../../../src/http/retry";
import { TeamsDevPortalClient } from "../../../src/clients/teamsDevPortal/client";
import { createMockContext } from "../testHelper";

/** Minimal buffer with valid ZIP local-file-header magic bytes (PK\x03\x04) */
const FAKE_ZIP = Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00]);

describe("TeamsDevPortalClient", () => {
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
    // Bypass retry delays — call fn() once, rethrow on error
    sandbox.stub(retryModule, "sendWithRetry").callsFake(async (fn: any) => fn());
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
  });

  afterEach(() => {
    sandbox.restore();
  });

  function createClient(): TeamsDevPortalClient {
    const ctx = createMockContext();
    return new TeamsDevPortalClient(ctx, "test-token-123");
  }

  describe("constructor", () => {
    it("sets Authorization and Client-Source headers", () => {
      createClient();
      expect(mockAxios.defaults.headers.common["Authorization"]).to.equal("Bearer test-token-123");
      expect(mockAxios.defaults.headers.common["Client-Source"]).to.equal("teamstoolkit");
    });
  });

  describe("importApp", () => {
    it("creates app with overwrite=false", async () => {
      const appDef = { teamsAppId: "app-1", tenantId: "tenant-1" };
      mockAxios.post.resolves({ data: appDef });

      const client = createClient();
      const result = await client.importApp(FAKE_ZIP, false);

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.deep.equal(appDef);
      expect(mockAxios.post.calledOnce).to.be.true;

      const [url, _body, opts] = mockAxios.post.firstCall.args;
      expect(url).to.equal("/api/appdefinitions/v2/import");
      expect(opts.params.overwriteIfAppAlreadyExists).to.be.false;
    });

    it("updates existing app with overwrite=true", async () => {
      const appDef = { teamsAppId: "app-1", tenantId: "tenant-1" };
      mockAxios.post.resolves({ data: appDef });

      const client = createClient();
      const result = await client.importApp(FAKE_ZIP, true);

      expect(result.isOk()).to.be.true;
      const [, , opts] = mockAxios.post.firstCall.args;
      expect(opts.params.overwriteIfAppAlreadyExists).to.be.true;
    });

    it("returns error on empty response", async () => {
      mockAxios.post.resolves({ data: null });

      const client = createClient();
      const result = await client.importApp(FAKE_ZIP);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsDevPortalImportError");
    });

    it("returns user error on 409 conflict", async () => {
      mockAxios.post.rejects({
        response: { status: 409, data: "conflict" },
        message: "Request failed with status code 409",
      });

      const client = createClient();
      const result = await client.importApp(FAKE_ZIP);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppConflictError");
      expect(result._unsafeUnwrapErr().kind).to.equal("user");
    });

    it("returns system error on 500", async () => {
      mockAxios.post.rejects({
        response: { status: 500, data: "server error" },
        message: "Internal server error",
      });

      const client = createClient();
      const result = await client.importApp(FAKE_ZIP);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });

  describe("getApp", () => {
    it("returns app definition when found", async () => {
      const appDef = { teamsAppId: "app-123", tenantId: "tenant-1" };
      mockAxios.get.resolves({ data: appDef });

      const client = createClient();
      const result = await client.getApp("app-123");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().teamsAppId).to.equal("app-123");
      expect(mockAxios.get.firstCall.args[0]).to.equal("/api/appdefinitions/app-123");
    });

    it("returns error when app ID mismatch", async () => {
      mockAxios.get.resolves({ data: { teamsAppId: "wrong-id" } });

      const client = createClient();
      const result = await client.getApp("app-123");

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppNotFound");
    });

    it("returns error on 404", async () => {
      mockAxios.get.rejects({
        response: { status: 404, data: "not found" },
        message: "Not found",
      });

      const client = createClient();
      const result = await client.getApp("app-123");

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppNotFound");
    });
  });

  describe("publishTeamsApp", () => {
    it("returns published app ID on success", async () => {
      mockAxios.post.resolves({ data: { id: "published-1" } });

      const client = createClient();
      const result = await client.publishTeamsApp("app-1", FAKE_ZIP);

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("published-1");
      expect(mockAxios.post.firstCall.args[0]).to.equal("/api/publishing");
    });

    it("returns error on API error response", async () => {
      mockAxios.post.resolves({ data: { error: { code: "SomeError", message: "fail" } } });

      const client = createClient();
      const result = await client.publishTeamsApp("app-1", FAKE_ZIP);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppPublishError");
    });

    it("falls through to update on Conflict/AppDefinitionAlreadyExists", async () => {
      // First call (publish) returns conflict
      mockAxios.post.onFirstCall().resolves({
        data: {
          error: {
            code: "Conflict",
            innerError: { code: "AppDefinitionAlreadyExists" },
          },
        },
      });
      // getStagedApp call returns staged app
      mockAxios.get.resolves({
        data: {
          value: [
            {
              appDefinitions: [
                { teamsAppId: "staged-1", displayName: "App", publishingState: "published" },
              ],
            },
          ],
        },
      });
      // Second call (update) succeeds
      mockAxios.post.onSecondCall().resolves({ data: { teamsAppId: "updated-1" } });

      const client = createClient();
      const result = await client.publishTeamsApp("app-1", FAKE_ZIP);

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("updated-1");
    });
  });

  describe("getStagedApp", () => {
    it("returns latest published definition", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              appDefinitions: [
                { teamsAppId: "old-1", displayName: "Old", publishingState: "submitted" },
                {
                  teamsAppId: "new-1",
                  displayName: "New",
                  publishingState: "published",
                  lastModifiedDateTime: "2026-03-30T00:00:00Z",
                },
              ],
            },
          ],
        },
      });

      const client = createClient();
      const result = await client.getStagedApp("app-1");

      expect(result.isOk()).to.be.true;
      const staged = result._unsafeUnwrap();
      expect(staged).to.not.be.undefined;
      expect(staged!.teamsAppId).to.equal("new-1");
      expect(staged!.publishingState).to.equal("published");
    });

    it("returns undefined when not published", async () => {
      mockAxios.get.resolves({ data: { value: [] } });

      const client = createClient();
      const result = await client.getStagedApp("app-1");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });

    it("returns undefined on HTTP error (graceful)", async () => {
      mockAxios.get.rejects(new Error("network error"));

      const client = createClient();
      const result = await client.getStagedApp("app-1");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });
  });

  describe("publishTeamsAppUpdate", () => {
    it("updates an already-published app", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              appDefinitions: [
                { teamsAppId: "staged-1", displayName: "App", publishingState: "published" },
              ],
            },
          ],
        },
      });
      mockAxios.post.resolves({ data: { teamsAppId: "updated-1" } });

      const client = createClient();
      const result = await client.publishTeamsAppUpdate("app-1", FAKE_ZIP);

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("updated-1");
      expect(mockAxios.post.firstCall.args[0]).to.include(
        "/api/publishing/staged-1/appdefinitions"
      );
    });

    it("returns error when app not published", async () => {
      mockAxios.get.resolves({ data: { value: [] } });

      const client = createClient();
      const result = await client.publishTeamsAppUpdate("app-1", FAKE_ZIP);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppNotPublished");
    });
  });

  // ── OAuth configuration methods ─────────────────────────────

  describe("getOauthRegistration", () => {
    it("returns OAuth config when found", async () => {
      mockAxios.get.resolves({
        data: { oAuthConfigId: "oauth-123", clientId: "client-abc" },
      });

      const client = createClient();
      const result = await client.getOauthRegistration("oauth-123");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()?.oAuthConfigId).to.equal("oauth-123");
      expect(mockAxios.get.firstCall.args[0]).to.include("/api/v1.0/oAuthConfigurations/oauth-123");
    });

    it("returns undefined when not found (404)", async () => {
      mockAxios.get.rejects({ response: { status: 404 } });

      const client = createClient();
      const result = await client.getOauthRegistration("missing");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });
  });

  describe("createOauthRegistration", () => {
    it("creates OAuth config and returns config ID", async () => {
      mockAxios.post.resolves({
        data: {
          configurationRegistrationId: { oAuthConfigId: "new-oauth-456" },
          resourceIdentifierUri: "api://example.com",
        },
      });

      const client = createClient();
      const result = await client.createOauthRegistration({
        description: "test",
        clientId: "client-abc",
        applicableToApps: "AnyApp" as any,
        targetUrlsShouldStartWith: ["https://example.com"],
      } as any);

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().configurationRegistrationId.oAuthConfigId).to.equal(
        "new-oauth-456"
      );
      expect(mockAxios.post.firstCall.args[0]).to.include("/api/v1.0/oAuthConfigurations");
    });
  });

  // ── API Key registration methods ────────────────────────────

  describe("getApiKeyRegistration", () => {
    it("returns API key registration when found", async () => {
      mockAxios.get.resolves({
        data: { id: "apikey-123", description: "test key" },
      });

      const client = createClient();
      const result = await client.getApiKeyRegistration("apikey-123");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()?.id).to.equal("apikey-123");
      expect(mockAxios.get.firstCall.args[0]).to.include(
        "/api/v1.0/apiSecretRegistrations/apikey-123"
      );
    });

    it("returns undefined when not found (404)", async () => {
      mockAxios.get.rejects({ response: { status: 404 } });

      const client = createClient();
      const result = await client.getApiKeyRegistration("missing");

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });
  });

  describe("createApiKeyRegistration", () => {
    it("creates API key registration and returns with ID", async () => {
      mockAxios.post.resolves({
        data: { id: "new-apikey-789", description: "test" },
      });

      const client = createClient();
      const result = await client.createApiKeyRegistration({
        description: "test",
        clientSecrets: [{ value: "secret12345", priority: 0 }],
        applicableToApps: "AnyApp" as any,
        targetUrlsShouldStartWith: ["https://api.example.com"],
      } as any);

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().id).to.equal("new-apikey-789");
      expect(mockAxios.post.firstCall.args[0]).to.include("/api/v1.0/apiSecretRegistrations");
    });
  });
});
