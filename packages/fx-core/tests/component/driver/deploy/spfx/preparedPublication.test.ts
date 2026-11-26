// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import * as fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { AppManifestUtils } from "@microsoft/teamsfx-api";

import {
  AppPackageOverrides,
  appId,
  createAppPackage,
  createPluginManifest,
  createSppkg,
  productId,
  publicationTenantContext,
} from "./preparedPublicationFixtures";
import { SHAREPOINT_PACKAGE_LIMITS } from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";
import {
  PreparedPublicationInput,
  preparePublication,
} from "../../../../../src/component/driver/deploy/spfx/utility/publicationArtifacts";

const sharePointInput = "sharepoint/solution.sppkg";
const applicationInput = "appPackage/application.zip";

function pluginWith(
  changes: Record<string, unknown>,
  runtime: Record<string, unknown> = {}
): Record<string, unknown> {
  return {
    ...createPluginManifest(),
    ...changes,
    runtimes: [
      {
        type: "RemoteMCPServer",
        auth: { type: "None" },
        spec: { url: "{{TENANT_MCP_URL}}" },
        ...runtime,
      },
    ],
  };
}

function nestedValue(depth: number): Record<string, unknown> {
  let value: Record<string, unknown> = { leaf: "value" };
  for (let level = 0; level < depth; level += 1) {
    value = { child: value };
  }
  return value;
}

function abortAfterChecks(checks: number): AbortSignal {
  const controller = new AbortController();
  let observed = 0;
  Object.defineProperty(controller.signal, "aborted", { get: () => observed++ >= checks });
  return controller.signal;
}

describe("SharePoint publication artifacts", () => {
  const directories: string[] = [];

  afterEach(async () => {
    vi.restoreAllMocks();
    await Promise.all(directories.splice(0).map((directory) => fs.remove(directory)));
  });

  async function createTemporaryDirectory(prefix: string): Promise<string> {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), prefix));
    directories.push(directory);
    return directory;
  }

  async function createProject(
    sharePointPackage = createSppkg(),
    applicationPackage = createAppPackage()
  ): Promise<string> {
    const projectPath = await createTemporaryDirectory("publication-project-");
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, sharePointInput), sharePointPackage);
    await fs.writeFile(path.join(projectPath, applicationInput), applicationPackage);
    return projectPath;
  }

  async function stagedEntries(projectPath: string): Promise<string[]> {
    const stagingRoot = path.join(projectPath, ".fx", "sharepoint-publications");
    return (await fs.pathExists(stagingRoot)) ? fs.readdir(stagingRoot) : [];
  }

  async function expectRejectedWithoutStaging(
    projectPath: string,
    code: string,
    overrides: Partial<PreparedPublicationInput> = {}
  ): Promise<void> {
    const sourceSharePoint = await fs.readFile(path.join(projectPath, sharePointInput));
    const sourceApplication = await fs.readFile(path.join(projectPath, applicationInput));

    const result = await preparePublication({
      projectPath,
      packagePath: sharePointInput,
      appPackagePath: applicationInput,
      tenantContext: publicationTenantContext,
      ...overrides,
    });

    expect(result.isErr() ? result.error.code : "Succeeded").toBe(code);
    expect(await stagedEntries(projectPath)).toEqual([]);
    expect(await fs.readFile(path.join(projectPath, sharePointInput))).toEqual(sourceSharePoint);
    expect(await fs.readFile(path.join(projectPath, applicationInput))).toEqual(sourceApplication);
  }

  it("SPFX-PREP-05: writes staged packages without a journal or source mutation", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "publication-artifacts-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    const sourceSharePoint = createSppkg();
    const sourceApplication = createAppPackage();
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), sourceSharePoint);
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), sourceApplication);

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    if (result.isErr()) return;
    expect(result.value).toEqual({
      sharePointPackagePath: expect.stringMatching(
        /^\.fx\/sharepoint-publications\/[0-9a-f-]+\/sharepoint\.sppkg$/
      ),
      appPackagePath: expect.stringMatching(
        /^\.fx\/sharepoint-publications\/[0-9a-f-]+\/application\.zip$/
      ),
    });
    expect(result.value.sharePointPackagePath).not.toBe(result.value.appPackagePath);
    expect(await fs.readFile(path.join(projectPath, "sharepoint", "solution.sppkg"))).toEqual(
      sourceSharePoint
    );
    expect(await fs.readFile(path.join(projectPath, "appPackage", "application.zip"))).toEqual(
      sourceApplication
    );

    const stagedSharePoint = new AdmZip(
      await fs.readFile(path.join(projectPath, result.value.sharePointPackagePath))
    );
    expect(stagedSharePoint.getEntry("ClientSideAssets/agent.zip")).toBeNull();
    expect(stagedSharePoint.readAsText("_rels/ClientSideAssets.xml.rels")).not.toContain(
      "agent.zip"
    );

    const stagedApplication = new AdmZip(
      await fs.readFile(path.join(projectPath, result.value.appPackagePath))
    );
    const manifest = JSON.parse(stagedApplication.readAsText("manifest.json"));
    expect(manifest.id).toBe(appId);
    const agent = JSON.parse(stagedApplication.readAsText("declarativeAgent.json"));
    expect(agent.instructions).toBe("Keep these instructions");
    const plugin = JSON.parse(stagedApplication.readAsText(agent.actions[0].file));
    expect(plugin.runtimes[0].spec.url).toBe(
      `https://tenant.example.test/_api/mcp/spfx?solutionId=${productId}`
    );

    const stagingDirectory = path.dirname(
      path.join(projectPath, result.value.sharePointPackagePath)
    );
    expect((await fs.readdir(stagingDirectory)).sort()).toEqual([
      "application.zip",
      "sharepoint.sppkg",
    ]);
  });

  it("SPFX-PREP-04: rejects invalid input paths without creating staging artifacts", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "publication-invalid-input-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());

    const result = await preparePublication({
      projectPath,
      packagePath: "../solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });

    expect(result.isErr()).toBe(true);
    if (result.isOk()) return;
    expect(result.error.code).toBe("InvalidInputPath");
    expect(await fs.pathExists(path.join(projectPath, ".fx"))).toBe(false);
  });

  it.each<[string, (projectPath: string) => Partial<PreparedPublicationInput>, string]>([
    [
      "an absolute package path",
      (projectPath) => ({ packagePath: path.join(projectPath, sharePointInput) }),
      "InvalidInputPath",
    ],
    ["a missing package", () => ({ packagePath: "sharepoint/missing.sppkg" }), "InputReadFailed"],
    ["a directory package path", () => ({ packagePath: "sharepoint" }), "InputReadFailed"],
    [
      "a missing application package",
      () => ({ appPackagePath: "appPackage/missing.zip" }),
      "InputReadFailed",
    ],
    [
      "a different selected SharePoint solution",
      () => ({
        tenantContext: {
          ...publicationTenantContext,
          productId: "22222222-3333-4444-8555-666666666666",
        },
      }),
      "SelectedProductMismatch",
    ],
  ])("SPFX-PREP-04: rejects %s before staging", async (_name, overrides, code) => {
    const projectPath = await createProject();

    await expectRejectedWithoutStaging(projectPath, code, overrides(projectPath));
  });

  it("SPFX-PREP-04: rejects an unsupported SharePoint package file name", async () => {
    const projectPath = await createProject();
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.zip"), createSppkg());

    await expectRejectedWithoutStaging(projectPath, "InvalidPackageName", {
      packagePath: "sharepoint/solution.zip",
    });
  });

  it("SPFX-PREP-04: rejects a linked input that resolves outside the project", async () => {
    const projectPath = await createProject();
    const outsidePath = await createTemporaryDirectory("publication-outside-");
    await fs.writeFile(path.join(outsidePath, "solution.sppkg"), createSppkg());
    await fs.symlink(outsidePath, path.join(projectPath, "sharepoint", "linked"), "junction");

    await expectRejectedWithoutStaging(projectPath, "InvalidInputPath", {
      packagePath: "sharepoint/linked/solution.sppkg",
    });
  });

  it("SPFX-PREP-04: rejects an input that exceeds the bounded read limit", async () => {
    const projectPath = await createProject();
    const largePath = path.join(projectPath, "sharepoint", "large.sppkg");
    await fs.writeFile(largePath, "");
    await fs.truncate(largePath, SHAREPOINT_PACKAGE_LIMITS.inputBytes + 1);

    await expectRejectedWithoutStaging(projectPath, "InputTooLarge", {
      packagePath: "sharepoint/large.sppkg",
    });
  });

  it("SPFX-PREP-04: rejects a SharePoint package that is not a valid archive", async () => {
    const projectPath = await createProject(Buffer.from("not an archive"));

    await expectRejectedWithoutStaging(projectPath, "InvalidArchive");
  });

  it("SPFX-PREP-04: rejects an unavailable project directory", async () => {
    const parentPath = await createTemporaryDirectory("publication-missing-project-");

    const result = await preparePublication({
      projectPath: path.join(parentPath, "missing"),
      packagePath: sharePointInput,
      appPackagePath: applicationInput,
      tenantContext: publicationTenantContext,
    });

    expect(result.isErr() ? result.error.code : "Succeeded").toBe("ProjectPathUnavailable");
    expect(await fs.readdir(parentPath)).toEqual([]);
  });

  it.each([
    ["an invalid tenant URL", { tenantUrl: "not a URL" }],
    ["a tenant id that is not a GUID", { tenantId: "tenant" }],
    ["a tenant URL without HTTPS", { tenantUrl: "http://tenant.example.test" }],
    ["a tenant URL with credentials", { tenantUrl: "https://user@tenant.example.test" }],
    ["a tenant URL with a query", { tenantUrl: "https://tenant.example.test/?site=root" }],
    ["a tenant URL with a fragment", { tenantUrl: "https://tenant.example.test/#root" }],
  ])("SPFX-PREP-09: rejects %s before staging", async (_name, tenantChanges) => {
    const projectPath = await createProject();

    await expectRejectedWithoutStaging(projectPath, "InvalidPreparedTenantContext", {
      tenantContext: { ...publicationTenantContext, ...tenantChanges },
    });
  });

  it.each<[string, Record<string, unknown>, string]>([
    [
      "an endpoint with another placeholder",
      pluginWith({}, { spec: { url: "{{OTHER_MCP_URL}}" } }),
      "UnresolvedPluginPlaceholder",
    ],
    [
      "a literal endpoint",
      pluginWith({}, { spec: { url: "https://mcp.example.test" } }),
      "AmbiguousPluginEndpoint",
    ],
    [
      "a runtime without a specification",
      pluginWith({}, { spec: undefined }),
      "AmbiguousPluginEndpoint",
    ],
    [
      "a content security entry with another placeholder",
      pluginWith(
        {},
        {
          spec: {
            url: "{{TENANT_MCP_URL}}",
            mcp_tool_description: {
              tools: [{ name: "tool", _meta: { "ui/csp": { frameDomains: ["{{CDN_ORIGIN}}"] } } }],
            },
          },
        }
      ),
      "UnresolvedPluginPlaceholder",
    ],
    [
      "another whole-value placeholder",
      pluginWith({ name_for_human: "{{PLUGIN_NAME}}" }),
      "UnresolvedPluginPlaceholder",
    ],
    [
      "a tenant origin outside the qualified content security lists",
      pluginWith({ description_for_human: "{{TENANT_ORIGIN}}" }),
      "UnqualifiedTenantOriginField",
    ],
    ["excessive nesting", pluginWith({ extension: nestedValue(130) }), "InvalidPluginManifest"],
  ])("SPFX-PREP-15: rejects a plugin with %s before staging", async (_name, plugin, code) => {
    const projectPath = await createProject(createSppkg("1.2.3.4", "component", plugin));

    await expectRejectedWithoutStaging(projectPath, code);
  });

  it.each<[string, AppPackageOverrides, string]>([
    [
      "an unsupported Teams manifest version",
      { manifest: { manifestVersion: "1.99" } },
      "UnsupportedTeamsManifestVersion",
    ],
    [
      "a Teams manifest that fails its schema",
      { manifest: { developer: undefined } },
      "ApplicationSchemaValidationFailed",
    ],
    [
      "an unsupported declarative agent version",
      { agent: { version: "v1.1" } },
      "UnsupportedAgentManifestVersion",
    ],
    [
      "a declarative agent that fails its schema",
      { agent: { description: undefined } },
      "ApplicationSchemaValidationFailed",
    ],
    [
      "an action outside the package",
      { agent: { actions: [{ id: "other", file: "../outside.json" }] } },
      "ApplicationSchemaValidationFailed",
    ],
    [
      "an action without a package entry",
      { agent: { actions: [{ id: "other", file: "missing.json" }] } },
      "ApplicationSchemaValidationFailed",
    ],
    [
      "an action plugin that fails its schema",
      {
        agent: { actions: [{ id: "other", file: "other.json" }] },
        files: { "other.json": JSON.stringify({ schema_version: "v2.4" }) },
      },
      "ApplicationSchemaValidationFailed",
    ],
  ])(
    "SPFX-PREP-14: rejects an application with %s before staging",
    async (_name, overrides, code) => {
      const projectPath = await createProject(createSppkg(), createAppPackage("2.3.4", overrides));

      await expectRejectedWithoutStaging(projectPath, code);
    }
  );

  it("SPFX-PREP-14: rejects an application whose schemas cannot be loaded", async () => {
    const projectPath = await createProject();
    vi.spyOn(AppManifestUtils, "fetchSchema").mockRejectedValue(new Error("schema unavailable"));

    await expectRejectedWithoutStaging(projectPath, "ApplicationSchemaValidationFailed");
  });

  it("SPFX-PREP-05: fails without staged output when the staging root cannot be created", async () => {
    const projectPath = await createProject();
    await fs.writeFile(path.join(projectPath, ".fx"), "not a directory");

    await expectRejectedWithoutStaging(projectPath, "PublicationPrepareFailed");
    expect(await fs.readFile(path.join(projectPath, ".fx"), "utf8")).toBe("not a directory");
  });

  it("SPFX-PREP-05: refuses a staging root that resolves outside the project", async () => {
    const projectPath = await createProject();
    const outsidePath = await createTemporaryDirectory("publication-staging-outside-");
    await fs.ensureDir(path.join(projectPath, ".fx"));
    await fs.symlink(
      outsidePath,
      path.join(projectPath, ".fx", "sharepoint-publications"),
      "junction"
    );

    await expectRejectedWithoutStaging(projectPath, "InvalidStagingPath");
    expect(await fs.readdir(outsidePath)).toEqual([]);
  });

  it.each([0, 1, 2, 3])(
    "SPFX-PREP-05: leaves no staged output when cancelled at checkpoint %i",
    async (checks) => {
      const projectPath = await createProject();

      await expectRejectedWithoutStaging(projectPath, "PublicationCancelled", {
        signal: abortAfterChecks(checks),
      });
    }
  );
});
