// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { createHash } from "crypto";
import fs from "fs-extra";
import os from "os";
import path from "path";
import YAML from "yaml";
import { afterEach, describe, expect, it, vi } from "vitest";

import * as Tools from "../../../../../src/common/tools";
import { SPFxDeployDriver } from "../../../../../src/component/driver/deploy/spfx/deployDriver";
import { inspectSharePointPackage } from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";
import {
  loadPreparedPublication,
  preparePublication,
  PreparedPublicationJournal,
} from "../../../../../src/component/driver/deploy/spfx/utility/preparedPublication";
import { MockTools } from "../../../../core/utils";
import { FxCoreClient } from "../../../../../src/core/FxCoreClient";
import { Validator } from "../../../../../src/component/configManager/validator";

import {
  appId,
  createAppPackage,
  createSppkg,
  mockTokenClaims,
  productId,
  publicationTenantContext,
  tenantId,
  tenantUrl,
} from "./preparedPublicationFixtures";
describe("Prepared SharePoint publication integration", () => {
  let projectPath: string | undefined;
  let http: MockAdapter | undefined;
  const adapters: MockAdapter[] = [];
  const originalTenantMcpUrl = process.env.TENANT_MCP_URL;

  afterEach(async () => {
    http?.restore();
    http = undefined;
    adapters.splice(0).forEach((adapter) => adapter.restore());
    vi.restoreAllMocks();
    if (originalTenantMcpUrl === undefined) {
      delete process.env.TENANT_MCP_URL;
    } else {
      process.env.TENANT_MCP_URL = originalTenantMcpUrl;
    }
    if (projectPath !== undefined) {
      await fs.remove(projectPath);
      projectPath = undefined;
    }
  });

  it.each([
    { outcome: "unknown" as const, makeUncertain: true },
    { outcome: "rejected" as const, makeUncertain: false },
  ])(
    "SPFX-PREP-08: records a $outcome upload and never replays it",
    async ({ outcome, makeUncertain }) => {
      projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-integration-"));
      await fs.ensureDir(path.join(projectPath, "sharepoint"));
      await fs.ensureDir(path.join(projectPath, "appPackage"));
      await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
      await fs.writeFile(
        path.join(projectPath, "appPackage", "application.zip"),
        createAppPackage()
      );
      process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
      const prepared = await preparePublication({
        projectPath,
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
        tenantContext: publicationTenantContext,
      });
      expect(prepared.isErr() ? prepared.error : undefined).toBeUndefined();
      if (prepared.isErr()) return;

      const requester = axios.create({ baseURL: tenantUrl });
      http = new MockAdapter(requester, { onNoMatch: "throwException" });
      vi.spyOn(axios, "create").mockReturnValue(requester);
      http.onGet("/_api/SP_TenantSettings_Current").reply(200, {
        CorporateCatalogUrl: `${tenantUrl}/sites/appcatalog`,
      });
      http.onGet(/lists\/getbytitle\('Apps for SharePoint'\)/).reply(200, { value: [] });
      http.onGet(/AvailableApps\/GetById/).reply(404);
      const upload = http.onPost(/tenantappcatalog\/Add\(overwrite=false,/);
      if (makeUncertain) {
        upload.networkErrorOnce();
      } else {
        upload.reply(403, { message: "SharePoint rejected the package upload." });
      }

      vi.spyOn(SPFxDeployDriver.prototype, "getTenant").mockResolvedValue(tenantUrl);
      vi.spyOn(Tools, "getSPFxToken").mockResolvedValue("sharepoint-token");
      const tools = new MockTools();
      mockTokenClaims(tools.tokenProvider.m365TokenProvider);
      vi.spyOn(tools.ui, "showMessage").mockImplementation(async () => ok("Confirm"));
      const context = {
        azureAccountProvider: tools.tokenProvider.azureAccountProvider,
        m365TokenProvider: tools.tokenProvider.m365TokenProvider,
        ui: tools.ui,
        progressBar: undefined,
        logProvider: tools.logProvider,
        projectPath,
        platform: Platform.VSCode,
      };
      const args = {
        createAppCatalogIfNotExist: false,
        packageSolutionPath: "sharepoint/config/package-solution.json",
        preparedPublicationPath: prepared.value.preparedPublicationPath,
        publicationId: prepared.value.publicationId,
        preparedPackagePath: prepared.value.sharePointPackagePath,
        sourcePackageSha256: prepared.value.sourcePackageSha256,
        sourceAppPackageSha256: prepared.value.sourceAppPackageSha256,
        pluginSha256: prepared.value.pluginSha256,
        sharePointPackageSha256: prepared.value.sharePointPackageSha256,
        appPackageSha256: prepared.value.appPackageSha256,
      };
      const driver = new SPFxDeployDriver();

      const first = await driver.run(args, context);
      expect(first.isErr() ? first.error.message : undefined).toBe(
        makeUncertain ? "SharePointUploadOutcomeUnknown" : "SharePointUploadRejected"
      );
      const loaded = await loadPreparedPublication(
        projectPath,
        prepared.value.publicationId,
        prepared.value.preparedPublicationPath
      );
      expect(loaded.isErr() ? loaded.error : undefined).toBeUndefined();
      if (loaded.isErr()) return;
      expect(loaded.value.journal.state).toBe(outcome);
      expect(loaded.value.journal.sharePointArtifact.uploadState).toBe(outcome);
      const writesAfterFirstAttempt = http.history.post.length;

      const replay = await driver.run(args, context);
      expect(replay.isErr()).toBe(true);
      expect(http.history.post).toHaveLength(writesAfterFirstAttempt);
      expect(http.history.post).toHaveLength(1);
      const freshPreparation = await preparePublication({
        projectPath,
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
        tenantContext: publicationTenantContext,
      });
      expect(freshPreparation.isErr()).toBe(true);
    }
  );

  it.each([
    {
      name: "same-version package with different bytes",
      version: "1.2.3.4",
      component: "different component",
      code: "SharePointPackageVersionCollision",
    },
    {
      name: "package older than the app catalog version",
      version: "1.2.3.5",
      component: "component",
      code: "SharePointPackageDowngrade",
    },
  ])(
    "SPFX-PREP-06: rejects $name before any SharePoint write",
    async ({ version, component, code }) => {
      projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-collision-"));
      await fs.ensureDir(path.join(projectPath, "sharepoint"));
      await fs.ensureDir(path.join(projectPath, "appPackage"));
      await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
      await fs.writeFile(
        path.join(projectPath, "appPackage", "application.zip"),
        createAppPackage()
      );
      process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
      const prepared = await preparePublication({
        projectPath,
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
        tenantContext: publicationTenantContext,
      });
      expect(prepared.isErr() ? prepared.error : undefined).toBeUndefined();
      if (prepared.isErr()) return;

      const existing = await inspectSharePointPackage(createSppkg(version, component));
      expect(existing.isErr() ? existing.error : undefined).toBeUndefined();
      if (existing.isErr()) return;
      const requester = axios.create({ baseURL: tenantUrl });
      http = new MockAdapter(requester, { onNoMatch: "throwException" });
      vi.spyOn(axios, "create").mockReturnValue(requester);
      http.onGet("/_api/SP_TenantSettings_Current").reply(200, {
        CorporateCatalogUrl: `${tenantUrl}/sites/appcatalog`,
      });
      http.onGet(/lists\/getbytitle\('Apps for SharePoint'\)/).reply(200, {
        value: [{ File: { ServerRelativeUrl: "/sites/appcatalog/solution.sppkg" } }],
      });
      http.onGet(/GetFileByServerRelativeUrl/).reply(200, existing.value.deploymentBytes);
      http.onGet(/AvailableApps\/GetById/).reply(200, {
        ID: productId,
        AppCatalogVersion: version,
        Deployed: true,
        CurrentVersionDeployed: true,
        IsPackageDefaultSkipFeatureDeployment: false,
      });

      vi.spyOn(SPFxDeployDriver.prototype, "getTenant").mockResolvedValue(tenantUrl);
      vi.spyOn(Tools, "getSPFxToken").mockResolvedValue("sharepoint-token");
      const tools = new MockTools();
      mockTokenClaims(tools.tokenProvider.m365TokenProvider);
      const approvalSpy = vi.spyOn(tools.ui, "showMessage");
      const context = {
        azureAccountProvider: tools.tokenProvider.azureAccountProvider,
        m365TokenProvider: tools.tokenProvider.m365TokenProvider,
        ui: tools.ui,
        progressBar: undefined,
        logProvider: tools.logProvider,
        projectPath,
        platform: Platform.VSCode,
      };
      const result = await new SPFxDeployDriver().run(
        {
          createAppCatalogIfNotExist: false,
          packageSolutionPath: "sharepoint/config/package-solution.json",
          preparedPublicationPath: prepared.value.preparedPublicationPath,
          publicationId: prepared.value.publicationId,
          preparedPackagePath: prepared.value.sharePointPackagePath,
          sourcePackageSha256: prepared.value.sourcePackageSha256,
          sourceAppPackageSha256: prepared.value.sourceAppPackageSha256,
          pluginSha256: prepared.value.pluginSha256,
          sharePointPackageSha256: prepared.value.sharePointPackageSha256,
          appPackageSha256: prepared.value.appPackageSha256,
        },
        context
      );

      expect(result.isErr() ? result.error.message : undefined).toBe(code);
      expect(approvalSpy).not.toHaveBeenCalled();
      expect(http.history.post).toHaveLength(0);
    }
  );

  it("SPFX-PREP-06: reuses identical current catalog content without another write", async () => {
    projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-reuse-"));
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
    const prepared = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });
    expect(prepared.isErr() ? prepared.error : undefined).toBeUndefined();
    if (prepared.isErr()) return;
    const packageBytes = await fs.readFile(
      path.join(projectPath, prepared.value.sharePointPackagePath)
    );

    const requester = axios.create({ baseURL: tenantUrl });
    http = new MockAdapter(requester, { onNoMatch: "throwException" });
    vi.spyOn(axios, "create").mockReturnValue(requester);
    http.onGet("/_api/SP_TenantSettings_Current").reply(200, {
      CorporateCatalogUrl: `${tenantUrl}/sites/appcatalog`,
    });
    http.onGet(/lists\/getbytitle\('Apps for SharePoint'\)/).reply(200, {
      value: [{ File: { ServerRelativeUrl: "/sites/appcatalog/solution.sppkg" } }],
    });
    http.onGet(/GetFileByServerRelativeUrl/).reply(200, packageBytes);
    http.onGet(/AvailableApps\/GetById/).reply(200, {
      ID: productId,
      AppCatalogVersion: "1.2.3.4",
      Deployed: true,
      CurrentVersionDeployed: true,
      IsPackageDefaultSkipFeatureDeployment: false,
    });
    vi.spyOn(SPFxDeployDriver.prototype, "getTenant").mockResolvedValue(tenantUrl);
    vi.spyOn(Tools, "getSPFxToken").mockResolvedValue("sharepoint-token");
    const tools = new MockTools();
    mockTokenClaims(tools.tokenProvider.m365TokenProvider);
    const approvalSpy = vi.spyOn(tools.ui, "showMessage");
    const context = {
      azureAccountProvider: tools.tokenProvider.azureAccountProvider,
      m365TokenProvider: tools.tokenProvider.m365TokenProvider,
      ui: tools.ui,
      progressBar: undefined,
      logProvider: tools.logProvider,
      projectPath,
      platform: Platform.VSCode,
    };
    const result = await new SPFxDeployDriver().run(
      {
        createAppCatalogIfNotExist: false,
        packageSolutionPath: "sharepoint/config/package-solution.json",
        preparedPublicationPath: prepared.value.preparedPublicationPath,
        publicationId: prepared.value.publicationId,
        preparedPackagePath: prepared.value.sharePointPackagePath,
        sourcePackageSha256: prepared.value.sourcePackageSha256,
        sourceAppPackageSha256: prepared.value.sourceAppPackageSha256,
        pluginSha256: prepared.value.pluginSha256,
        sharePointPackageSha256: prepared.value.sharePointPackageSha256,
        appPackageSha256: prepared.value.appPackageSha256,
      },
      context
    );
    const journal = await loadPreparedPublication(
      projectPath,
      prepared.value.publicationId,
      prepared.value.preparedPublicationPath
    );

    expect(
      result.isOk(),
      result.isErr()
        ? `${result.error.name}: ${result.error.message}; data=${JSON.stringify(result.error.data)}`
        : ""
    ).toBe(true);
    expect(journal.isErr() ? journal.error : undefined).toBeUndefined();
    if (journal.isOk()) {
      expect(journal.value.journal.state).toBe("sharePointVerified");
      expect(journal.value.journal.sharePointArtifact.uploadState).toBe("confirmed");
      expect(journal.value.journal.sharePointArtifact.deployState).toBe("confirmed");
    }
    expect(approvalSpy).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
  });

  it("SPFX-PREP-09: runs the YAML publish lifecycle SharePoint-first with hash-bound artifacts", async () => {
    projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-publish-client-"));
    await fs.ensureDir(path.join(projectPath, "env"));
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(
      path.join(projectPath, "env", ".env.dev"),
      "TEAMSFX_ENV=dev\nTENANT_MCP_URL=https://tenant.example.test/mcp\n"
    );
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    await fs.writeFile(
      path.join(projectPath, "m365agents.yml"),
      [
        "version: v1.14",
        "environmentFolderPath: ./env",
        "publish:",
        "  - uses: spfx/preparePublication",
        "    with:",
        "      packagePath: ./sharepoint/solution.sppkg",
        "      appPackagePath: ./appPackage/application.zip",
        "    writeToEnvironmentFile:",
        "      publicationId: PREP_PUBLICATION_ID",
        "      preparedPublicationPath: PREP_PUBLICATION_PATH",
        "      sourcePackageSha256: PREP_SOURCE_PACKAGE_SHA256",
        "      sourceAppPackageSha256: PREP_SOURCE_APP_PACKAGE_SHA256",
        "      pluginSha256: PREP_PLUGIN_SHA256",
        "      sharePointPackagePath: PREP_SHAREPOINT_PACKAGE_PATH",
        "      sharePointPackageSha256: PREP_SHAREPOINT_PACKAGE_SHA256",
        "      appPackagePath: PREP_APP_PACKAGE_PATH",
        "      appPackageSha256: PREP_APP_PACKAGE_SHA256",
        "  - uses: spfx/deploy",
        "    with:",
        "      createAppCatalogIfNotExist: false",
        "      packageSolutionPath: ./sharepoint/config/package-solution.json",
        "      preparedPackagePath: ${{PREP_SHAREPOINT_PACKAGE_PATH}}",
        "      preparedPublicationPath: ${{PREP_PUBLICATION_PATH}}",
        "      publicationId: ${{PREP_PUBLICATION_ID}}",
        "      sourcePackageSha256: ${{PREP_SOURCE_PACKAGE_SHA256}}",
        "      sourceAppPackageSha256: ${{PREP_SOURCE_APP_PACKAGE_SHA256}}",
        "      pluginSha256: ${{PREP_PLUGIN_SHA256}}",
        "      sharePointPackageSha256: ${{PREP_SHAREPOINT_PACKAGE_SHA256}}",
        "      appPackageSha256: ${{PREP_APP_PACKAGE_SHA256}}",
        "  - uses: teamsApp/publishAppPackage",
        "    with:",
        "      appPackagePath: ${{PREP_APP_PACKAGE_PATH}}",
        "      preparedPublicationPath: ${{PREP_PUBLICATION_PATH}}",
        "      publicationId: ${{PREP_PUBLICATION_ID}}",
        "      sourcePackageSha256: ${{PREP_SOURCE_PACKAGE_SHA256}}",
        "      sourceAppPackageSha256: ${{PREP_SOURCE_APP_PACKAGE_SHA256}}",
        "      pluginSha256: ${{PREP_PLUGIN_SHA256}}",
        "      sharePointPackageSha256: ${{PREP_SHAREPOINT_PACKAGE_SHA256}}",
        "      appPackageSha256: ${{PREP_APP_PACKAGE_SHA256}}",
        "    writeToEnvironmentFile:",
        "      publishedAppId: M365_PUBLISHED_APP_ID",
        "",
      ].join("\n")
    );
    const yamlContent = YAML.parse(
      await fs.readFile(path.join(projectPath, "m365agents.yml"), "utf8")
    );
    const validator = new Validator().impl.get("v1.14")?.validator;
    expect(validator?.(yamlContent), JSON.stringify(validator?.errors)).toBe(true);

    process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
    const tools = new MockTools();
    vi.spyOn(tools.tokenProvider.m365TokenProvider, "getJsonObject").mockImplementation(
      async (tokenRequest) =>
        ok({
          tid: tenantId,
          aud: tokenRequest.scopes.some((scope) => scope.includes("Sites.FullControl.All"))
            ? tenantUrl
            : "https://graph.microsoft.com",
        })
    );
    let cancelFirstApproval = true;
    let abortAtNextApproval = false;
    const approvalAbortSignal = new AbortController();
    vi.spyOn(tools.ui, "showMessage").mockImplementation(
      async (_level, _message, modal, ...items) => {
        if (!modal) return ok("");
        if (cancelFirstApproval) {
          cancelFirstApproval = false;
          return ok("Cancel");
        }
        if (abortAtNextApproval) {
          abortAtNextApproval = false;
          approvalAbortSignal?.abort();
        }
        return ok(items[0]);
      }
    );
    vi.spyOn(SPFxDeployDriver.prototype, "getTenant").mockResolvedValue(tenantUrl);
    vi.spyOn(Tools, "getSPFxToken").mockResolvedValue("sharepoint-token");

    let uploaded = false;
    let deployed = false;
    let catalogVersion = "1.2.3.4";
    let preparedCatalogVersion = "1.2.3.4";
    let graphAppCreated = false;
    const graphAppVersion = "2.3.4";
    let abortNextSharePointUpload = false;
    const uploadSignal = new AbortController();
    let abortNextSharePointDeploy = false;
    const deploySignal = new AbortController();
    let failNextGraphSubmission = false;
    let abortNextGraphSubmission = false;
    let submissionSignal: AbortController | undefined;
    let uploadedPackage: Buffer | undefined;
    let submittedPackage: Buffer | undefined;
    const writes: string[] = [];
    const createAxiosInstance = axios.create.bind(axios);
    vi.spyOn(axios, "create").mockImplementation((config) => {
      const requester = createAxiosInstance(config);
      const adapter = new MockAdapter(requester, { onNoMatch: "throwException" });
      adapters.push(adapter);
      adapter.onAny().reply((request) => {
        const url = request.url ?? "";
        if (request.method === "get" && url.includes("SP_TenantSettings_Current")) {
          return [200, { CorporateCatalogUrl: `${tenantUrl}/sites/appcatalog` }];
        }
        if (request.method === "get" && url.includes("lists/getbytitle")) {
          return [
            200,
            {
              value: uploaded
                ? [{ File: { ServerRelativeUrl: "/sites/appcatalog/solution.sppkg" } }]
                : [],
            },
          ];
        }
        if (request.method === "get" && url.includes("GetFileByServerRelativeUrl")) {
          return [200, uploadedPackage ?? Buffer.alloc(0)];
        }
        if (request.method === "get" && url.includes("AvailableApps/GetById")) {
          if (!uploaded) return [404, {}];
          return [
            200,
            {
              ID: productId,
              AppCatalogVersion: catalogVersion,
              Deployed: deployed,
              CurrentVersionDeployed: deployed,
              IsPackageDefaultSkipFeatureDeployment: false,
            },
          ];
        }
        if (request.method === "post" && url.includes("tenantappcatalog/Add(")) {
          if (abortNextSharePointUpload) {
            abortNextSharePointUpload = false;
            uploadSignal?.abort();
          }
          uploaded = true;
          deployed = false;
          catalogVersion = preparedCatalogVersion;
          uploadedPackage = Buffer.isBuffer(request.data)
            ? request.data
            : Buffer.from(String(request.data));
          writes.push("sharepoint-upload");
          return [200, {}];
        }
        if (
          request.method === "post" &&
          url.includes("AvailableApps/GetById") &&
          url.includes("/Deploy")
        ) {
          if (abortNextSharePointDeploy) {
            abortNextSharePointDeploy = false;
            deploySignal?.abort();
          }
          deployed = true;
          writes.push("sharepoint-deploy");
          return [200, {}];
        }
        if (request.method === "get" && url.includes("/appCatalogs/teamsApps")) {
          return [
            200,
            {
              value: graphAppCreated
                ? [
                    {
                      externalId: appId,
                      id: "published-teams-app-id",
                      displayName: "Prepared application",
                      appDefinitions: [{ version: graphAppVersion, publishingState: "published" }],
                    },
                  ]
                : [],
            },
          ];
        }
        if (request.method === "post" && url.includes("/appCatalogs/teamsApps")) {
          submittedPackage = Buffer.isBuffer(request.data)
            ? request.data
            : Buffer.from(String(request.data));
          writes.push("teams-app-submit");
          if (abortNextGraphSubmission) {
            abortNextGraphSubmission = false;
            submissionSignal?.abort();
          }
          if (failNextGraphSubmission) {
            failNextGraphSubmission = false;
            return [500, { message: "Submission outcome unknown" }];
          }
          graphAppCreated = true;
          return [201, { externalId: appId, id: "published-teams-app-id" }];
        }
        return [500, { message: `Unexpected HTTP request: ${request.method} ${url}` }];
      });
      return requester;
    });

    const client = new FxCoreClient(tools);
    const alreadyAborted = new AbortController();
    alreadyAborted.abort();
    const cancelledBeforeStart = await client.publish(
      {
        platform: Platform.VSCode,
        nonInteractive: true,
        projectPath,
        env: "dev",
      },
      { signal: alreadyAborted.signal }
    );
    expect(cancelledBeforeStart.isErr()).toBe(true);
    expect(writes).toEqual([]);

    const cancelled = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });
    expect(cancelled.isErr()).toBe(true);
    expect(writes).toEqual([]);
    const publicationRoot = path.join(projectPath, ".fx", "sharepoint-publications");
    expect(
      await fs.pathExists(publicationRoot),
      cancelled.isErr() ? `${cancelled.error.name}: ${cancelled.error.message}` : ""
    ).toBe(true);
    const cancelledRecords = (await fs.readdir(publicationRoot, { withFileTypes: true })).filter(
      (record) => record.isDirectory() && /^[0-9a-f-]{36}$/i.test(record.name)
    );
    expect(cancelledRecords).toHaveLength(1);
    const cancelledJournalPath = path
      .join(".fx", "sharepoint-publications", cancelledRecords[0].name, "journal.json")
      .split(path.sep)
      .join("/");
    const cancelledJournal = await loadPreparedPublication(
      projectPath,
      cancelledRecords[0].name,
      cancelledJournalPath
    );
    expect(cancelledJournal.isErr() ? cancelledJournal.error : undefined).toBeUndefined();
    if (cancelledJournal.isOk()) {
      expect(cancelledJournal.value.journal.state).toBe("prepared");
    }

    abortAtNextApproval = true;
    const cancelledBeforeUpload = await client.publish(
      {
        platform: Platform.VSCode,
        nonInteractive: true,
        projectPath,
        env: "dev",
      },
      { signal: approvalAbortSignal.signal }
    );
    expect(cancelledBeforeUpload.isErr()).toBe(true);
    expect(writes).toEqual([]);
    const cancelledBeforeUploadJournal = await loadPreparedPublication(
      projectPath,
      cancelledRecords[0].name,
      cancelledJournalPath
    );
    expect(
      cancelledBeforeUploadJournal.isErr() ? cancelledBeforeUploadJournal.error : undefined
    ).toBeUndefined();
    if (cancelledBeforeUploadJournal.isOk()) {
      expect(cancelledBeforeUploadJournal.value.journal.state).toBe("prepared");
      expect(cancelledBeforeUploadJournal.value.journal.sharePointArtifact.uploadState).toBe(
        "notStarted"
      );
    }

    abortNextSharePointUpload = true;
    const cancelledDuringUpload = await client.publish(
      {
        platform: Platform.VSCode,
        nonInteractive: true,
        projectPath,
        env: "dev",
      },
      { signal: uploadSignal.signal }
    );
    expect(cancelledDuringUpload.isErr()).toBe(true);
    expect(writes).toEqual(["sharepoint-upload"]);
    const uploadedBeforeCancellation = await loadPreparedPublication(
      projectPath,
      cancelledRecords[0].name,
      cancelledJournalPath
    );
    expect(
      uploadedBeforeCancellation.isErr() ? uploadedBeforeCancellation.error : undefined
    ).toBeUndefined();
    if (uploadedBeforeCancellation.isOk()) {
      expect(uploadedBeforeCancellation.value.journal.state).toBe("sharePointUploaded");
      expect(uploadedBeforeCancellation.value.journal.sharePointArtifact.uploadState).toBe(
        "confirmed"
      );
    }

    const result = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });

    expect(
      result.isOk(),
      result.isErr()
        ? `${result.error.name}: ${result.error.message}; data=${JSON.stringify(result.error.data)}`
        : ""
    ).toBe(true);
    expect(writes).toEqual(["sharepoint-upload", "sharepoint-deploy", "teams-app-submit"]);
    expect(uploadedPackage).toBeDefined();
    expect(submittedPackage).toBeDefined();
    const outputs = result.isOk() ? result.value.outputs : {};
    const journal = await loadPreparedPublication(
      projectPath,
      outputs["PREP_PUBLICATION_ID"] ?? "",
      outputs["PREP_PUBLICATION_PATH"] ?? ""
    );
    expect(journal.isErr() ? journal.error : undefined).toBeUndefined();
    if (journal.isOk()) {
      expect(journal.value.journal.state).toBe("submitted");
      expect(journal.value.journal.sharePointArtifact.deployState).toBe("confirmed");
      expect(
        createHash("sha256")
          .update(uploadedPackage ?? Buffer.alloc(0))
          .digest("hex")
      ).toBe(journal.value.journal.sharePointArtifact.sha256);
      expect(
        createHash("sha256")
          .update(submittedPackage ?? Buffer.alloc(0))
          .digest("hex")
      ).toBe(journal.value.journal.applicationArtifact.sha256);
      expect(journal.value.journal.applicationArtifact.submittedAppId).toBe(
        "published-teams-app-id"
      );
    }

    const writesAfterSuccessfulPublication = writes.length;
    const successfulReplay = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });
    expect(
      successfulReplay.isOk(),
      successfulReplay.isErr()
        ? `${successfulReplay.error.name}: ${successfulReplay.error.message}; data=${JSON.stringify(successfulReplay.error.data)}`
        : ""
    ).toBe(true);
    expect(writes).toHaveLength(writesAfterSuccessfulPublication);
    if (successfulReplay.isOk() && result.isOk()) {
      expect(successfulReplay.value.outputs["PREP_PUBLICATION_ID"]).toBe(
        result.value.outputs["PREP_PUBLICATION_ID"]
      );
      expect(successfulReplay.value.outputs["M365_PUBLISHED_APP_ID"]).toBe(
        "published-teams-app-id"
      );
    }

    await fs.writeFile(
      path.join(projectPath, "sharepoint", "solution.sppkg"),
      createSppkg("1.2.3.5")
    );
    await fs.writeFile(
      path.join(projectPath, "appPackage", "application.zip"),
      createAppPackage("2.3.5")
    );
    preparedCatalogVersion = "1.2.3.5";
    abortNextSharePointDeploy = true;
    const writesBeforeCancelledDeploy = writes.length;
    const cancelledDuringDeploy = await client.publish(
      {
        platform: Platform.VSCode,
        nonInteractive: true,
        projectPath,
        env: "dev",
      },
      { signal: deploySignal.signal }
    );
    expect(cancelledDuringDeploy.isErr()).toBe(true);
    expect(writes.slice(writesBeforeCancelledDeploy)).toEqual([
      "sharepoint-upload",
      "sharepoint-deploy",
    ]);
    const writesBeforeDeployResume = writes.length;
    const resumedAfterDeployCancellation = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });
    expect(
      resumedAfterDeployCancellation.isOk(),
      resumedAfterDeployCancellation.isErr()
        ? `${resumedAfterDeployCancellation.error.name}: ${resumedAfterDeployCancellation.error.message}`
        : ""
    ).toBe(true);
    expect(writes.slice(writesBeforeDeployResume)).toEqual(["teams-app-submit"]);

    await fs.writeFile(
      path.join(projectPath, "sharepoint", "solution.sppkg"),
      createSppkg("1.2.3.6")
    );
    await fs.writeFile(
      path.join(projectPath, "appPackage", "application.zip"),
      createAppPackage("2.3.6")
    );
    preparedCatalogVersion = "1.2.3.6";
    submissionSignal = new AbortController();
    abortNextGraphSubmission = true;
    const writesBeforeCancelledSubmission = writes.length;
    const cancelledDuringSubmission = await client.publish(
      {
        platform: Platform.VSCode,
        nonInteractive: true,
        projectPath,
        env: "dev",
      },
      { signal: submissionSignal.signal }
    );
    expect(cancelledDuringSubmission.isErr()).toBe(true);
    expect(writes.slice(writesBeforeCancelledSubmission)).toEqual([
      "sharepoint-upload",
      "sharepoint-deploy",
      "teams-app-submit",
    ]);
    const settledRecords = await fs.readdir(publicationRoot, { withFileTypes: true });
    const settledSubmissionJournals: PreparedPublicationJournal[] = [];
    for (const record of settledRecords.filter(
      (entry) => entry.isDirectory() && /^[0-9a-f-]{36}$/i.test(entry.name)
    )) {
      const journalPath = path
        .join(".fx", "sharepoint-publications", record.name, "journal.json")
        .split(path.sep)
        .join("/");
      const loaded = await loadPreparedPublication(projectPath, record.name, journalPath);
      if (loaded.isOk() && loaded.value.journal.sourceSharePoint.version === "1.2.3.6") {
        settledSubmissionJournals.push(loaded.value.journal);
      }
    }
    expect(settledSubmissionJournals).toHaveLength(1);
    expect(settledSubmissionJournals[0].state).toBe("submitted");
    expect(settledSubmissionJournals[0].applicationArtifact.submissionState).toBe("confirmed");

    await fs.writeFile(
      path.join(projectPath, "sharepoint", "solution.sppkg"),
      createSppkg("1.2.3.7")
    );
    await fs.writeFile(
      path.join(projectPath, "appPackage", "application.zip"),
      createAppPackage("2.3.7")
    );
    preparedCatalogVersion = "1.2.3.7";
    failNextGraphSubmission = true;
    submissionSignal = new AbortController();
    abortNextGraphSubmission = true;
    const writesBeforeUnknownOutcome = writes.length;
    const unknownOutcome = await client.publish(
      {
        platform: Platform.VSCode,
        nonInteractive: true,
        projectPath,
        env: "dev",
      },
      { signal: submissionSignal.signal }
    );
    expect(unknownOutcome.isErr()).toBe(true);
    if (unknownOutcome.isErr()) {
      expect(unknownOutcome.error.data).toMatchObject({
        stage: "applicationSubmission",
        lastDurableState: "unknown",
        lastSharePointVerifiedState: "sharePointVerified",
        httpStatus: 500,
      });
      expect(unknownOutcome.error.data).not.toHaveProperty("response");
    }
    expect(writes.slice(writesBeforeUnknownOutcome)).toEqual([
      "sharepoint-upload",
      "sharepoint-deploy",
      "teams-app-submit",
    ]);

    const records = await fs.readdir(publicationRoot, { withFileTypes: true });
    const unresolvedRecords: PreparedPublicationJournal[] = [];
    for (const record of records.filter(
      (entry) => entry.isDirectory() && /^[0-9a-f-]{36}$/i.test(entry.name)
    )) {
      const journalPath = path
        .join(".fx", "sharepoint-publications", record.name, "journal.json")
        .split(path.sep)
        .join("/");
      const loaded = await loadPreparedPublication(projectPath, record.name, journalPath);
      if (loaded.isOk() && loaded.value.journal.state === "unknown") {
        unresolvedRecords.push(loaded.value.journal);
      }
    }
    expect(unresolvedRecords).toHaveLength(1);
    expect(unresolvedRecords[0].applicationArtifact.submissionState).toBe("unknown");
    const writesAfterUnknownOutcome = writes.length;
    const replay = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });
    expect(replay.isErr()).toBe(true);
    expect(writes).toHaveLength(writesAfterUnknownOutcome);
  });
});
