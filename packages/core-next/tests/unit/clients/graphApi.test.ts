/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import axios from "axios";
import * as retryModule from "../../../src/http/retry";
import { createMockContext } from "../testHelper";
import { GraphApiClient } from "../../../src/clients/graphApi/client";
import { AADApplication } from "../../../src/clients/graphApi/types";

describe("GraphApiClient", () => {
  let sandbox: sinon.SinonSandbox;
  let mockAxios: {
    post: sinon.SinonStub;
    get: sinon.SinonStub;
    patch: sinon.SinonStub;
    delete: sinon.SinonStub;
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
      delete: sandbox.stub(),
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

  describe("constructor", () => {
    it("sets Authorization and Content-Type headers", () => {
      const ctx = createMockContext();
      new GraphApiClient(ctx, "test-graph-token");
      expect(mockAxios.defaults.headers.common["Authorization"]).to.equal(
        "Bearer test-graph-token"
      );
      expect(mockAxios.defaults.headers.common["Content-Type"]).to.equal("application/json");
    });
  });

  describe("createAadApp", () => {
    it("creates an app with displayName", async () => {
      const app: AADApplication = { id: "obj-1", appId: "client-1", displayName: "My App" };
      mockAxios.post.resolves({ data: app, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.createAadApp("My App");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().appId).to.equal("client-1");
      expect(mockAxios.post.calledOnce).to.be.true;
      expect(mockAxios.post.firstCall.args[0]).to.equal("/applications");
    });

    it("passes signInAudience and serviceManagementReference", async () => {
      mockAxios.post.resolves({ data: { id: "o1", appId: "c1" }, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      await client.createAadApp("App", "AzureADMultipleOrgs", "svc-ref-123");
      const body = mockAxios.post.firstCall.args[1];
      expect(body.signInAudience).to.equal("AzureADMultipleOrgs");
      expect(body.serviceManagementReference).to.equal("svc-ref-123");
    });

    it("returns user error on 4xx", async () => {
      mockAxios.post.rejects({
        response: {
          status: 400,
          data: { error: { code: "BadRequest", message: "Invalid audience" } },
        },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.createAadApp("App");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("user");
    });

    it("returns system error on 5xx", async () => {
      mockAxios.post.rejects({
        response: { status: 500, data: { error: { code: "InternalError", message: "boom" } } },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.createAadApp("App");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });

  describe("generateClientSecret", () => {
    it("returns secretText on success", async () => {
      mockAxios.post.resolves({
        data: { secretText: "super-secret-value", keyId: "k1" },
        status: 200,
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.generateClientSecret("obj-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("super-secret-value");
    });

    it("returns error on empty secretText", async () => {
      mockAxios.post.resolves({ data: { secretText: "" }, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.generateClientSecret("obj-123");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("EmptyClientSecret");
    });

    it("uses extra retries", async () => {
      mockAxios.post.resolves({ data: { secretText: "s" }, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");
      await client.generateClientSecret("obj-123");

      const retryStub = retryModule.sendWithRetry as sinon.SinonStub;
      // second arg should be 5
      expect(retryStub.calledOnce).to.be.true;
      expect(retryStub.firstCall.args[1]).to.equal(5);
    });
  });

  describe("updateAadApp", () => {
    it("PATCHes the application", async () => {
      mockAxios.patch.resolves({ data: {}, status: 204 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.updateAadApp("obj-123", { displayName: "New Name" });
      expect(result.isOk()).to.be.true;
      expect(mockAxios.patch.firstCall.args[0]).to.equal("/applications/obj-123");
    });

    it("uses extra retries", async () => {
      mockAxios.patch.resolves({ data: {}, status: 204 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");
      await client.updateAadApp("obj-123", {});

      const retryStub = retryModule.sendWithRetry as sinon.SinonStub;
      expect(retryStub.firstCall.args[1]).to.equal(5);
    });
  });

  describe("getOwners", () => {
    it("returns owner list", async () => {
      mockAxios.get.resolves({
        data: { value: [{ id: "u1", displayName: "User 1" }] },
        status: 200,
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getOwners("obj-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.have.lengthOf(1);
    });

    it("returns empty array when value is missing", async () => {
      mockAxios.get.resolves({ data: {}, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getOwners("obj-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.deep.equal([]);
    });
  });

  describe("addOwner", () => {
    it("posts $ref link", async () => {
      mockAxios.post.resolves({ data: {}, status: 204 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.addOwner("obj-123", "user-456");
      expect(result.isOk()).to.be.true;
      expect(mockAxios.post.firstCall.args[0]).to.equal("/applications/obj-123/owners/$ref");
    });
  });

  /* ------------------------------------------------------------------ */
  /*  App Catalog — publish operations                                   */
  /* ------------------------------------------------------------------ */

  describe("getStagedApp", () => {
    it("returns published app definition when found", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "catalog-id-1",
              displayName: "My App",
              appDefinitions: [
                {
                  lastModifiedDateTime: "2024-01-01T00:00:00Z",
                  publishingState: "published",
                },
              ],
            },
          ],
        },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getStagedApp("ext-id-1");
      expect(result.isOk()).to.be.true;
      const staged = result._unsafeUnwrap();
      expect(staged).to.not.be.undefined;
      expect(staged!.teamsAppId).to.equal("catalog-id-1");
      expect(staged!.displayName).to.equal("My App");
      expect(staged!.publishingState).to.equal("published");
      expect(staged!.lastModifiedDateTime).to.be.instanceOf(Date);
    });

    it("returns undefined when no apps match", async () => {
      mockAxios.get.resolves({ data: { value: [] } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getStagedApp("ext-id-1");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });

    it("returns undefined when appDefinitions is empty", async () => {
      mockAxios.get.resolves({ data: { value: [{ id: "x", appDefinitions: [] }] } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getStagedApp("ext-id-1");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });

    it("swallows errors and returns undefined", async () => {
      mockAxios.get.rejects(new Error("network failure"));
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getStagedApp("ext-id-1");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });

    it("returns latest appDefinition when multiple exist", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-1",
              displayName: "App",
              appDefinitions: [
                { publishingState: "submitted", lastModifiedDateTime: "2024-01-01T00:00:00Z" },
                { publishingState: "published", lastModifiedDateTime: "2024-06-01T00:00:00Z" },
              ],
            },
          ],
        },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getStagedApp("ext-id-1");
      expect(result._unsafeUnwrap()!.publishingState).to.equal("published");
    });

    it("uses the beta endpoint path", async () => {
      mockAxios.get.resolves({ data: { value: [] } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      await client.getStagedApp("ext-id-1");
      const url = mockAxios.get.firstCall.args[0] as string;
      expect(url).to.include("/appCatalogs/teamsApps");
      expect(url).to.include("externalId eq 'ext-id-1'");
    });
  });

  describe("publishTeamsApp", () => {
    const zipBuffer = Buffer.from("PK-fake-zip");

    it("returns app ID on successful publish", async () => {
      mockAxios.post.resolves({ data: { id: "published-id-1" }, status: 201 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-id-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("published-id-1");
    });

    it("sends ZIP with correct content type", async () => {
      mockAxios.post.resolves({ data: { id: "pub-1" } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      await client.publishTeamsApp("ext-1", zipBuffer);
      const callArgs = mockAxios.post.firstCall.args;
      expect(callArgs[0]).to.include("/appCatalogs/teamsApps?requiresReview=true");
      expect(callArgs[2].headers["Content-Type"]).to.equal("application/zip");
    });

    it("falls back to getStagedApp on BadGateway response body", async () => {
      mockAxios.post.resolves({
        data: { error: { code: "BadGateway", message: "bad gateway" } },
      });
      // getStagedApp will be called via the same mockAxios.get
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "fallback-id",
              displayName: "App",
              appDefinitions: [{ publishingState: "submitted", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("fallback-id");
    });

    it("falls through to update on Conflict/AppDefinitionAlreadyExists", async () => {
      // First call: publishTeamsApp POST → Conflict
      mockAxios.post.onFirstCall().resolves({
        data: {
          error: {
            code: "Conflict",
            innerError: { code: "AppDefinitionAlreadyExists" },
          },
        },
      });
      // getStagedApp → found
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "existing-cat-id",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      // publishTeamsAppUpdate POST → success
      mockAxios.post.onSecondCall().resolves({ data: { teamsAppId: "updated-id" } });

      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("updated-id");
    });

    it("falls through to update on HTTP 409", async () => {
      // POST throws 409
      mockAxios.post.onFirstCall().rejects({ response: { status: 409 } });
      // getStagedApp → found
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-409",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      // publishTeamsAppUpdate POST → success
      mockAxios.post.onSecondCall().resolves({ data: { id: "updated-409" } });

      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("updated-409");
    });

    it("falls back to getStagedApp when response has no id", async () => {
      mockAxios.post.resolves({ data: {} });
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "staged-fallback",
              displayName: "App",
              appDefinitions: [{ publishingState: "submitted", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("staged-fallback");
    });

    it("returns error when response empty and no staged app", async () => {
      mockAxios.post.resolves({ data: {} });
      mockAxios.get.resolves({ data: { value: [] } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-1", zipBuffer);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("GraphPublishError");
    });

    it("returns system error on non-409 HTTP errors", async () => {
      mockAxios.post.rejects({
        response: { status: 500, data: { error: { code: "InternalError", message: "boom" } } },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsApp("ext-1", zipBuffer);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });

  describe("publishTeamsAppUpdate", () => {
    const zipBuffer = Buffer.from("PK-fake-zip");

    it("updates published app and returns teamsAppId", async () => {
      // getStagedApp → found
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-upd-1",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      mockAxios.post.resolves({ data: { teamsAppId: "updated-id-1" } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("updated-id-1");
    });

    it("posts to correct appDefinitions endpoint", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-path",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      mockAxios.post.resolves({ data: { id: "x" } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      const url = mockAxios.post.firstCall.args[0] as string;
      expect(url).to.include("/appCatalogs/teamsApps/cat-path/appDefinitions?requiresReview=true");
    });

    it("returns error when app not found in catalog", async () => {
      mockAxios.get.resolves({ data: { value: [] } });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("TeamsAppNotPublished");
      expect(result._unsafeUnwrapErr().kind).to.equal("user");
    });

    it("falls back to staged teamsAppId when response has no id", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-fb",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      mockAxios.post.resolves({ data: {} });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("cat-fb");
    });

    it("returns error on Graph API error body", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-err",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      mockAxios.post.resolves({
        data: { error: { code: "SomethingBad", message: "error details" } },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("GraphPublishUpdateError");
    });

    it("returns system error on network failure", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-net",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      mockAxios.post.rejects(new Error("ECONNRESET"));
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });

    it("falls back without requiresReview on 400 (sideloaded app)", async () => {
      const stagedData = {
        data: {
          value: [
            {
              id: "cat-sideloaded",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      };
      mockAxios.get.resolves(stagedData);
      // First POST (with requiresReview) → 400
      const error400 = Object.assign(new Error("Bad Request"), {
        response: { status: 400, data: { error: { code: "BadRequest", message: "Bad Request" } } },
      });
      mockAxios.post.onFirstCall().rejects(error400);
      // Second POST (without requiresReview) → success
      mockAxios.post.onSecondCall().resolves({ data: { teamsAppId: "fallback-ok" } });

      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("fallback-ok");
      // First POST should have requiresReview, second should not
      const firstUrl = mockAxios.post.firstCall.args[0] as string;
      const secondUrl = mockAxios.post.secondCall.args[0] as string;
      expect(firstUrl).to.include("?requiresReview=true");
      expect(secondUrl).to.not.include("requiresReview");
    });

    it("returns existing catalog ID when 400 fallback also fails with 404 (sideloaded phantom)", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-phantom",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      const error400 = Object.assign(new Error("Bad Request"), {
        response: { status: 400, data: {} },
      });
      const error404 = Object.assign(new Error("Not Found"), {
        response: { status: 404, data: { error: { message: "App doesn't exist" } } },
      });
      mockAxios.post.onFirstCall().rejects(error400);
      mockAxios.post.onSecondCall().rejects(error404);

      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("cat-phantom");
    });

    it("returns error when 400 fallback fails with non-404 error", async () => {
      mockAxios.get.resolves({
        data: {
          value: [
            {
              id: "cat-fb-fail",
              displayName: "App",
              appDefinitions: [{ publishingState: "published", lastModifiedDateTime: null }],
            },
          ],
        },
      });
      const error400 = Object.assign(new Error("Bad Request"), {
        response: { status: 400, data: {} },
      });
      mockAxios.post.onFirstCall().rejects(error400);
      mockAxios.post.onSecondCall().rejects(new Error("Also failed"));

      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.publishTeamsAppUpdate("ext-1", zipBuffer);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });

  describe("unpublishTeamsApp", () => {
    it("deletes successfully and returns ok(undefined)", async () => {
      mockAxios.delete.resolves({ status: 204, data: "" });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.unpublishTeamsApp("catalog-id-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.be.undefined;
    });

    it("sends DELETE to correct appCatalogs path", async () => {
      mockAxios.delete.resolves({ status: 204, data: "" });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      await client.unpublishTeamsApp("catalog-id-456");
      const url = mockAxios.delete.firstCall.args[0] as string;
      expect(url).to.include("/appCatalogs/teamsApps/catalog-id-456");
    });

    it("returns error on 404 (app not found)", async () => {
      const err404 = Object.assign(new Error("Not Found"), {
        response: { status: 404, data: { error: { code: "NotFound", message: "App not found" } } },
      });
      mockAxios.delete.rejects(err404);
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.unpublishTeamsApp("nonexistent-id");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("user");
    });

    it("returns system error on network failure", async () => {
      mockAxios.delete.rejects(new Error("ECONNRESET"));
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.unpublishTeamsApp("catalog-id-789");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });
});
