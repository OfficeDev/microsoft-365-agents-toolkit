// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { Platform, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { createHash } from "crypto";
import fs from "fs-extra";
import os from "os";
import path from "path";
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

const productId = "11111111-2222-4333-8444-555555555555";
const appId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";
const componentId = "bbbbbbbb-cccc-4ddd-8eee-ffffffffffff";
const tenantUrl = "https://tenant.example.test";

function createSppkg(version = "1.2.3.4", componentContent = "component"): Buffer {
  const agent = new AdmZip();
  agent.addFile(
    "ai-plugin.json",
    Buffer.from(
      JSON.stringify({
        schema_version: "v2.4",
        namespace: "sample",
        functions: [],
        runtimes: [{ type: "RemoteMCPServer", spec: { url: "{{TENANT_MCP_URL}}" } }],
      })
    )
  );
  agent.addFile("manifest.json", Buffer.from("{}"));
  agent.addFile("declarativeAgent.json", Buffer.from("{}"));

  const packageZip = new AdmZip();
  packageZip.addFile(
    "AppManifest.xml",
    Buffer.from(
      `<App ProductID="{${productId}}" Version="${version}" IsClientSideSolution="true" SkipFeatureDeployment="false" />`
    )
  );
  packageZip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="zip" ContentType="application/octet-stream" /></Types>`
    )
  );
  packageZip.addFile(
    "_rels/.rels",
    Buffer.from(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="manifest" Target="/AppManifest.xml" Id="r0" /></Relationships>`
    )
  );
  packageZip.addFile(
    "_rels/ClientSideAssets.xml.rels",
    Buffer.from(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="agent" Target="/ClientSideAssets/agent.zip" Id="agent" /><Relationship Type="component" Target="/ClientSideAssets/component.js" Id="component" /></Relationships>`
    )
  );
  packageZip.addFile("ClientSideAssets.xml", Buffer.from("<Assets />"));
  packageZip.addFile(
    `${componentId}/CopilotComponent_${componentId}.xml`,
    Buffer.from("<Elements><ClientSideComponent /></Elements>")
  );
  packageZip.addFile("ClientSideAssets/agent.zip", agent.toBuffer());
  packageZip.addFile("ClientSideAssets/component.js", Buffer.from(componentContent));
  return packageZip.toBuffer();
}

function createAppPackage(version = "2.3.4"): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    "manifest.json",
    Buffer.from(
      JSON.stringify({
        manifestVersion: "1.30",
        version,
        id: appId,
        copilotAgents: {
          declarativeAgents: [{ id: "agent", file: "declarativeAgent.json" }],
        },
      })
    )
  );
  zip.addFile(
    "declarativeAgent.json",
    Buffer.from(JSON.stringify({ instructions: "Keep these instructions", actions: [] }))
  );
  return zip.toBuffer();
}

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
      });
      expect(prepared.isOk()).toBe(true);
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
      expect(first.isErr()).toBe(true);
      const loaded = await loadPreparedPublication(
        projectPath,
        prepared.value.publicationId,
        prepared.value.preparedPublicationPath
      );
      expect(loaded.isOk()).toBe(true);
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
      });
      expect(freshPreparation.isErr()).toBe(true);
    }
  );

  it.each([
    {
      name: "same-version package with different bytes",
      version: "1.2.3.4",
      component: "different component",
    },
    {
      name: "package older than the app catalog version",
      version: "1.2.3.5",
      component: "component",
    },
  ])("SPFX-PREP-06: rejects $name before any SharePoint write", async ({ version, component }) => {
    projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-collision-"));
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
    const prepared = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
    });
    expect(prepared.isOk()).toBe(true);
    if (prepared.isErr()) return;

    const existing = await inspectSharePointPackage(createSppkg(version, component));
    expect(existing.isOk()).toBe(true);
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

    expect(result.isErr()).toBe(true);
    expect(approvalSpy).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
  });

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
    });
    expect(prepared.isOk()).toBe(true);
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

    expect(result.isOk()).toBe(true);
    expect(journal.isOk()).toBe(true);
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

    process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
    const tools = new MockTools();
    let cancelFirstApproval = true;
    vi.spyOn(tools.ui, "showMessage").mockImplementation(
      async (_level, _message, _modal, ...items) => {
        if (cancelFirstApproval) {
          cancelFirstApproval = false;
          return ok("Cancel");
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
    let failNextGraphSubmission = false;
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
    const cancelled = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });
    expect(cancelled.isErr()).toBe(true);
    expect(writes).toEqual([]);
    const publicationRoot = path.join(projectPath, ".fx", "sharepoint-publications");
    const cancelledRecords = await fs.readdir(publicationRoot, { withFileTypes: true });
    expect(cancelledRecords.filter((record) => record.isDirectory())).toHaveLength(1);
    const cancelledJournalPath = path
      .join(".fx", "sharepoint-publications", cancelledRecords[0].name, "journal.json")
      .split(path.sep)
      .join("/");
    const cancelledJournal = await loadPreparedPublication(
      projectPath,
      cancelledRecords[0].name,
      cancelledJournalPath
    );
    expect(cancelledJournal.isOk()).toBe(true);
    if (cancelledJournal.isOk()) {
      expect(cancelledJournal.value.journal.state).toBe("prepared");
    }

    const result = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });

    expect(result.isOk()).toBe(true);
    expect(writes).toEqual(["sharepoint-upload", "sharepoint-deploy", "teams-app-submit"]);
    expect(uploadedPackage).toBeDefined();
    expect(submittedPackage).toBeDefined();
    const outputs = result.isOk() ? result.value.outputs : {};
    const journal = await loadPreparedPublication(
      projectPath,
      outputs["PREP_PUBLICATION_ID"] ?? "",
      outputs["PREP_PUBLICATION_PATH"] ?? ""
    );
    expect(journal.isOk()).toBe(true);
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

    await fs.writeFile(
      path.join(projectPath, "sharepoint", "solution.sppkg"),
      createSppkg("1.2.3.5")
    );
    await fs.writeFile(
      path.join(projectPath, "appPackage", "application.zip"),
      createAppPackage("2.3.5")
    );
    preparedCatalogVersion = "1.2.3.5";
    failNextGraphSubmission = true;
    const unknownOutcome = await client.publish({
      platform: Platform.VSCode,
      nonInteractive: true,
      projectPath,
      env: "dev",
    });
    expect(unknownOutcome.isErr()).toBe(true);
    expect(writes.slice(3)).toEqual(["sharepoint-upload", "sharepoint-deploy", "teams-app-submit"]);

    const records = await fs.readdir(publicationRoot, { withFileTypes: true });
    const unresolvedRecords: PreparedPublicationJournal[] = [];
    for (const record of records.filter((entry) => entry.isDirectory())) {
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
