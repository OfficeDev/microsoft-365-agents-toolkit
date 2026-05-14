/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs/promises";
import axios from "axios";
import AdmZip from "adm-zip";
import * as retryModule from "../../../src/http/retry";
import { createMockContext } from "../testHelper";
import { M365PackageService, AppScope } from "../../../src/clients/m365";

describe("M365PackageService", () => {
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
  let tmpDir: string;

  beforeEach(async () => {
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
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-test-"));
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  function createService(): M365PackageService {
    const ctx = createMockContext();
    return new M365PackageService(ctx, "mock-token");
  }

  /** Create a classic Teams app package (no copilotAgents). */
  function createClassicPackage(): string {
    const zip = new AdmZip();
    zip.addFile(
      "manifest.json",
      Buffer.from(JSON.stringify({ id: "test-app-id", name: { short: "App" } }))
    );
    const zipPath = path.join(tmpDir, "classic.zip");
    zip.writeZip(zipPath);
    return zipPath;
  }

  /** Create a Declarative Agent app package (has copilotAgents). */
  function createDAPackage(): string {
    const zip = new AdmZip();
    zip.addFile(
      "manifest.json",
      Buffer.from(
        JSON.stringify({
          id: "da-app-id",
          name: { short: "DA App" },
          copilotAgents: { declarativeAgents: [{ id: "agent1" }] },
        })
      )
    );
    const zipPath = path.join(tmpDir, "da.zip");
    zip.writeZip(zipPath);
    return zipPath;
  }

  describe("constructor", () => {
    it("sets Authorization header from token", () => {
      const svc = createService();
      expect(mockAxios.defaults.headers.common["Authorization"]).to.equal("Bearer mock-token");
      expect(svc).to.be.instanceOf(M365PackageService);
    });
  });

  describe("sideLoad — V1 (classic Teams app)", () => {
    it("returns titleId and appId on success", async () => {
      const zipPath = createClassicPackage();
      const svc = createService();

      // getTitleServiceUrl
      mockAxios.get.onFirstCall().resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // upload
      mockAxios.post.onFirstCall().resolves({ data: { operationId: "op-1" } });
      // acquire
      mockAxios.post.onSecondCall().resolves({ data: { statusId: "status-1" } });
      // poll — immediately ready
      mockAxios.get.onSecondCall().resolves({
        status: 200,
        data: { titleId: "title-abc", appId: "app-xyz" },
      });

      const result = await svc.sideLoad(zipPath);

      expect(result.isOk()).to.be.true;
      const val = result._unsafeUnwrap();
      expect(val.titleId).to.equal("title-abc");
      expect(val.appId).to.equal("app-xyz");
      expect(val.shareLink).to.equal("");
    });

    it("returns error for invalid ZIP without manifest.json", async () => {
      const zip = new AdmZip();
      zip.addFile("readme.txt", Buffer.from("hello"));
      const zipPath = path.join(tmpDir, "no-manifest.zip");
      zip.writeZip(zipPath);

      const svc = createService();
      const result = await svc.sideLoad(zipPath);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("InvalidAppPackage");
    });

    it("returns system error on HTTP 500", async () => {
      const zipPath = createClassicPackage();
      const svc = createService();

      // getTitleServiceUrl
      mockAxios.get.onFirstCall().resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // upload fails with 500
      const err500 = new Error("Server Error") as Error & {
        response: { status: number; data: string };
      };
      err500.response = { status: 500, data: "Internal Server Error" };
      mockAxios.post.onFirstCall().rejects(err500);

      const result = await svc.sideLoad(zipPath);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("M365ServiceError");
    });

    it("returns user error on HTTP 400", async () => {
      const zipPath = createClassicPackage();
      const svc = createService();

      // getTitleServiceUrl
      mockAxios.get.onFirstCall().resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // upload fails with 400
      const err400 = new Error("Bad Request") as Error & {
        response: { status: number; data: string };
      };
      err400.response = { status: 400, data: "Invalid package" };
      mockAxios.post.onFirstCall().rejects(err400);

      const result = await svc.sideLoad(zipPath);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("M365ServiceUserError");
    });
  });

  describe("sideLoad — V2 (Declarative Agent)", () => {
    it("returns titleId and appId via Builder API", async () => {
      const zipPath = createDAPackage();
      const svc = createService();

      // getTitleServiceUrl
      mockAxios.get.onFirstCall().resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // builder upload
      mockAxios.post.resolves({ data: { statusId: "builder-status-1" } });
      // poll — immediately ready
      mockAxios.get.onSecondCall().resolves({
        status: 200,
        data: { titleId: "da-title", appId: "da-app" },
      });

      const result = await svc.sideLoad(zipPath);

      expect(result.isOk()).to.be.true;
      const val = result._unsafeUnwrap();
      expect(val.titleId).to.equal("da-title");
      expect(val.appId).to.equal("da-app");
      expect(val.shareLink).to.equal("");
    });

    it("fetches shareLink when scope is Shared", async () => {
      const zipPath = createDAPackage();
      const svc = createService();

      // getTitleServiceUrl (called twice: once for upload, once for shareLink)
      mockAxios.get.onFirstCall().resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // builder upload
      mockAxios.post.resolves({ data: { statusId: "builder-status-2" } });
      // poll — immediately ready
      mockAxios.get.onSecondCall().resolves({
        status: 200,
        data: { titleId: "da-title-shared", appId: "da-app-shared" },
      });
      // getTitleServiceUrl for shareLink
      mockAxios.get.onCall(2).resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // shareLink response
      mockAxios.get.onCall(3).resolves({
        data: { shareLink: "https://example.com/share/12345" },
      });

      const result = await svc.sideLoad(zipPath, AppScope.Shared);

      expect(result.isOk()).to.be.true;
      const val = result._unsafeUnwrap();
      expect(val.titleId).to.equal("da-title-shared");
      expect(val.shareLink).to.equal("https://example.com/share/12345");
    });

    it("returns error when builder upload fails", async () => {
      const zipPath = createDAPackage();
      const svc = createService();

      // getTitleServiceUrl
      mockAxios.get.onFirstCall().resolves({
        data: { titlesServiceUrl: "https://titles.example.com" },
      });
      // builder upload fails
      mockAxios.post.rejects(new Error("Network error"));

      const result = await svc.sideLoad(zipPath);

      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("M365ServiceError");
    });
  });
});
