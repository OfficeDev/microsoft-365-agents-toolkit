// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  compareSharePointPackageVersions,
  createPreparedPublicationErrorData,
  isAmbiguousPublicationState,
  loadPreparedPublication,
  preparePublication,
  PreparedPublicationJournal,
  updatePreparedPublication,
  PreparedPublicationError,
} from "../../../../../src/component/driver/deploy/spfx/utility/preparedPublication";

const appId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";
const productId = "11111111-2222-4333-8444-555555555555";
const componentId = "bbbbbbbb-cccc-4ddd-8eee-ffffffffffff";
const tenantContext = {
  tenantId: "cccccccc-dddd-4eee-8fff-000000000000",
  tenantUrl: "https://tenant.example.test",
  catalogUrl: "https://tenant.example.test/sites/appcatalog",
  productId,
};

function createSppkg(
  options: { includeCsp?: boolean; includeUnrelatedPlaceholder?: boolean } = {}
): Buffer {
  const agent = new AdmZip();
  agent.addFile(
    "ai-plugin.json",
    Buffer.from(
      JSON.stringify({
        schema_version: "v2.4",
        namespace: "sample",
        name_for_human: "Sample",
        description_for_human: "Sample API plugin",
        description_for_model: options.includeUnrelatedPlaceholder
          ? "Literal text {{TENANT_MCP_URL}} must remain unchanged."
          : "Sample model description.",
        functions: [],
        runtimes: [
          {
            type: "RemoteMCPServer",
            auth: { type: "None" },
            spec: {
              url: "{{TENANT_MCP_URL}}",
              ...(options.includeCsp === true
                ? { content_security_policy: { connect_domains: ["{{TENANT_ORIGIN}}"] } }
                : {}),
            },
          },
        ],
      })
    )
  );
  agent.addFile("manifest.json", Buffer.from("{}"));
  agent.addFile("declarativeAgent.json", Buffer.from("{}"));

  const packageZip = new AdmZip();
  packageZip.addFile(
    "AppManifest.xml",
    Buffer.from(
      `<App ProductID="{${productId}}" Version="1.2.3.4" IsClientSideSolution="true" SkipFeatureDeployment="false" />`
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
  packageZip.addFile("ClientSideAssets/component.js", Buffer.from("component"));
  return packageZip.toBuffer();
}

function createAppPackage(connectorOnly = false): Buffer {
  const zip = new AdmZip();
  const manifest: Record<string, unknown> = {
    manifestVersion: connectorOnly ? "devPreview" : "1.30",
    version: "2.3.4",
    id: appId,
    ...(connectorOnly
      ? {
          $schema:
            "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
          developer: {
            name: "Sample developer",
            websiteUrl: "https://example.test",
            privacyUrl: "https://example.test/privacy",
            termsOfUseUrl: "https://example.test/terms",
          },
          name: { short: "Sample app", full: "Sample application" },
          description: { short: "Sample app", full: "Sample application" },
          icons: { color: "color.png", outline: "outline.png" },
          accentColor: "#FFFFFF",
        }
      : {
          $schema:
            "https://developer.microsoft.com/json-schemas/teams/v1.30/MicrosoftTeams.schema.json",
          developer: {
            name: "Sample developer",
            websiteUrl: "https://example.test",
            privacyUrl: "https://example.test/privacy",
            termsOfUseUrl: "https://example.test/terms",
          },
          name: { short: "Sample app", full: "Sample application" },
          description: { short: "Sample app", full: "Sample application" },
          icons: { color: "color.png", outline: "outline.png" },
          accentColor: "#FFFFFF",
        }),
    ...(connectorOnly
      ? {
          agentConnectors: [
            {
              id: "existing-connector",
              displayName: "Existing connector",
              toolSource: {
                remoteMcpServer: { mcpServerUrl: "https://existing.example.test/mcp" },
              },
            },
          ],
          agentSkills: [{ folder: "skills/existing" }],
        }
      : {
          copilotAgents: {
            declarativeAgents: [{ id: "agent", file: "declarativeAgent.json" }],
          },
        }),
  };
  zip.addFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
  if (!connectorOnly) {
    zip.addFile(
      "declarativeAgent.json",
      Buffer.from(
        JSON.stringify({
          version: "v1.9",
          name: "Sample agent",
          description: "Sample agent description",
          instructions: "Keep these instructions",
          actions: [],
        })
      )
    );
  }
  zip.addFile("instruction.txt", Buffer.from("Original application instructions"));
  zip.addFile("color.png", Buffer.from("icon"));
  zip.addFile("outline.png", Buffer.from("icon"));
  return zip.toBuffer();
}

describe("Prepared SharePoint publication preparation", () => {
  const directories: string[] = [];
  const originalTenantUrl = process.env.TENANT_MCP_URL;

  afterEach(async () => {
    if (originalTenantUrl === undefined) {
      delete process.env.TENANT_MCP_URL;
    } else {
      process.env.TENANT_MCP_URL = originalTenantUrl;
    }
    await Promise.all(directories.splice(0).map((directory) => fs.remove(directory)));
  });

  it("SPFX-PREP-06: compares SharePoint package versions without accepting malformed values", () => {
    expect(compareSharePointPackageVersions("1.2.3.4", "1.2.3.5")).toBe(-1);
    expect(compareSharePointPackageVersions("2.0.0.0", "1.99.99.99")).toBe(1);
    expect(compareSharePointPackageVersions("1.2.3.4", "1.2.3.4")).toBe(0);
    expect(compareSharePointPackageVersions("1.2.3", "1.2.3.4")).toBeUndefined();
  });

  it("SPFX-PREP-04: rejects a publication input that changes during its bounded read", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-publication-race-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    const sourcePath = path.join(projectPath, "sharepoint", "solution.sppkg");
    await fs.writeFile(sourcePath, createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    const originalRead = fs.read.bind(fs);
    let changed = false;
    const readSpy = vi
      .spyOn(fs, "read")
      .mockImplementation(async (fd, buffer, offset, length, position) => {
        const result = await originalRead(fd, buffer, offset, length, position);
        if (!changed) {
          changed = true;
          await fs.appendFile(sourcePath, Buffer.from("changed during read"));
        }
        return result;
      });

    try {
      const result = await preparePublication({
        projectPath,
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
        tenantContext,
      });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.code).toBe("InputChangedDuringRead");
      }
    } finally {
      readSpy.mockRestore();
    }
  });

  it("SPFX-PREP-11: projects rejected and unknown Graph failures without raw response data", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-publication-errors-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    const prepared = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext,
    });
    expect(prepared.isOk()).toBe(true);
    if (prepared.isErr()) return;
    const loaded = await loadPreparedPublication(
      projectPath,
      prepared.value.publicationId,
      prepared.value.preparedPublicationPath
    );
    expect(loaded.isOk()).toBe(true);
    if (loaded.isErr()) return;

    const rejectedJournal: PreparedPublicationJournal = {
      ...loaded.value.journal,
      state: "rejected",
      sharePointArtifact: {
        ...loaded.value.journal.sharePointArtifact,
        deployState: "confirmed",
      },
    };
    const rejectedResponse = {
      status: 400,
      data: { error: { code: "BadRequest", message: "do not expose response text" } },
      headers: { "request-id": "request-123" },
    };
    const rejectedCause = Object.assign(new Error("raw service response"), {
      response: rejectedResponse,
    });
    const rejected = createPreparedPublicationErrorData(
      new PreparedPublicationError("TeamsAppSubmissionRejected", "Rejected", {
        cause: rejectedCause,
      }),
      rejectedJournal,
      prepared.value.preparedPublicationPath,
      "applicationSubmission"
    );
    expect(rejected).toMatchObject({
      stage: "applicationSubmission",
      lastDurableState: "rejected",
      lastSharePointVerifiedState: "sharePointVerified",
      httpStatus: 400,
      serviceCode: "BadRequest",
      requestId: "request-123",
    });
    expect(rejected).not.toHaveProperty("response");
    expect(JSON.stringify(rejected)).not.toContain("do not expose response text");

    const unknownJournal: PreparedPublicationJournal = {
      ...rejectedJournal,
      state: "unknown",
      applicationArtifact: {
        ...rejectedJournal.applicationArtifact,
        submissionState: "unknown",
      },
    };
    const unknown = createPreparedPublicationErrorData(
      new PreparedPublicationError("TeamsAppSubmissionOutcomeUnknown", "Unknown", {
        cause: Object.assign(new Error("raw service response"), {
          response: { status: 500, data: { message: "do not expose response text" } },
        }),
      }),
      unknownJournal,
      prepared.value.preparedPublicationPath,
      "applicationSubmission"
    );
    expect(unknown).toMatchObject({
      lastDurableState: "unknown",
      lastSharePointVerifiedState: "sharePointVerified",
      httpStatus: 500,
    });
    expect(JSON.stringify(unknown)).not.toContain("do not expose response text");
  });

  it("SPFX-PREP-02/05: freezes and journals the derived artifacts without changing sources", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-publication-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    const sharePointSource = createSppkg({ includeUnrelatedPlaceholder: true });
    const appSource = createAppPackage();
    const sharePointPath = path.join(projectPath, "sharepoint", "solution.sppkg");
    const appPath = path.join(projectPath, "appPackage", "application.zip");
    await fs.writeFile(sharePointPath, sharePointSource);
    await fs.writeFile(appPath, appSource);
    process.env.TENANT_MCP_URL = "https://attacker.example.test/arbitrary";

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext,
    });

    expect(
      result.isOk(),
      result.isErr()
        ? `${result.error.code}: ${result.error.message}; cause=${String(result.error.cause)}`
        : ""
    ).toBe(true);
    if (result.isErr()) return;
    expect(await fs.readFile(sharePointPath)).toEqual(sharePointSource);
    expect(await fs.readFile(appPath)).toEqual(appSource);
    const loaded = await loadPreparedPublication(
      projectPath,
      result.value.publicationId,
      result.value.preparedPublicationPath
    );
    expect(loaded.isOk()).toBe(true);
    if (loaded.isErr()) return;
    expect(loaded.value.journal.state).toBe("prepared");
    expect(loaded.value.journal.sourceSharePoint.productId).toBe(productId);
    expect(loaded.value.journal.sourceApplication.appId).toBe(appId);
    expect(loaded.value.journal.identity.tenantOrigin).toBe(tenantContext.tenantUrl);
    expect(loaded.value.journal.identity.catalogUrl).toBe(tenantContext.catalogUrl);
    expect(JSON.stringify(loaded.value.journal)).not.toContain("attacker.example.test");

    const applicationPath = path.join(projectPath, result.value.appPackagePath);
    const stagedApp = new AdmZip(await fs.readFile(applicationPath));
    const stagedAgent = JSON.parse(stagedApp.readAsText("declarativeAgent.json"));
    expect(stagedAgent.instructions).toBe("Keep these instructions");
    expect(stagedAgent.actions).toHaveLength(1);
    const pluginPath = stagedAgent.actions[0].file;
    const stagedPlugin = JSON.parse(stagedApp.readAsText(pluginPath));
    expect(stagedPlugin.runtimes[0].spec.url).toBe(
      `https://tenant.example.test/_api/mcp/spfx?solutionId=${productId}`
    );
    expect(stagedPlugin.description_for_model).toBe(
      "Literal text {{TENANT_MCP_URL}} must remain unchanged."
    );

    const markedUnknown = await updatePreparedPublication(
      projectPath,
      result.value.publicationId,
      result.value.preparedPublicationPath,
      (journal) => ({ ...journal, state: "unknown" })
    );
    expect(markedUnknown.isOk()).toBe(true);
    if (markedUnknown.isOk()) {
      expect(isAmbiguousPublicationState(markedUnknown.value)).toBe(true);
    }
    const retry = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext,
    });
    expect(retry.isErr()).toBe(true);
    if (retry.isErr()) {
      expect(retry.error.code).toBe("UnresolvedPreparedPublication");
    }

    const changedSource = new AdmZip(createSppkg());
    changedSource.addFile("ClientSideAssets/changed.txt", Buffer.from("changed source"));
    await fs.writeFile(sharePointPath, changedSource.toBuffer());
    const changedSourceRetry = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext,
    });
    expect(changedSourceRetry.isErr()).toBe(true);
  });

  it("SPFX-PREP-02: supports connector-only apps without synthesizing an agent", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-publication-connector-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(
      path.join(projectPath, "appPackage", "application.zip"),
      createAppPackage(true)
    );
    process.env.TENANT_MCP_URL = "https://attacker.example.test/arbitrary";

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext,
    });

    expect(
      result.isOk(),
      result.isErr()
        ? `${result.error.code}: ${result.error.message}; cause=${String(result.error.cause)}`
        : ""
    ).toBe(true);
    if (result.isErr()) return;
    const stagedApp = new AdmZip(
      await fs.readFile(path.join(projectPath, result.value.appPackagePath))
    );
    const manifest = JSON.parse(stagedApp.readAsText("manifest.json"));
    expect(manifest.copilotAgents).toBeUndefined();
    expect(manifest.agentConnectors[0].id).toBe("existing-connector");
    expect(manifest.agentSkills[0].folder).toBe("skills/existing");
    expect(manifest.agentConnectors).toHaveLength(2);
    const preparedConnector = manifest.agentConnectors.find((connector: { id: string }) =>
      connector.id.startsWith("spfx-api-")
    );
    expect(preparedConnector.toolSource.plugin.file).toMatch(/^plugins\/spfx-api-/);
    expect(preparedConnector.toolSource.plugin.id).toBe("sample");
    expect(stagedApp.getEntry("declarativeAgent.json")).toBeNull();
  });

  it("SPFX-PREP-05: produces identical application bytes when system time changes", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-publication-stable-"));
    const secondProjectPath = await fs.mkdtemp(
      path.join(os.tmpdir(), "prepared-publication-stable-second-")
    );
    directories.push(projectPath);
    directories.push(secondProjectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.ensureDir(path.join(secondProjectPath, "sharepoint"));
    await fs.ensureDir(path.join(secondProjectPath, "appPackage"));
    const sharePointSource = createSppkg();
    const appSource = createAppPackage();
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), sharePointSource);
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), appSource);
    await fs.writeFile(
      path.join(secondProjectPath, "sharepoint", "solution.sppkg"),
      sharePointSource
    );
    await fs.writeFile(path.join(secondProjectPath, "appPackage", "application.zip"), appSource);
    process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";
    vi.useFakeTimers({ toFake: ["Date"] });
    try {
      vi.setSystemTime(new Date("2026-01-01T00:00:00.000Z"));
      const first = await preparePublication({
        projectPath,
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
        tenantContext,
      });
      expect(first.isOk(), first.isErr() ? `${first.error.code}: ${first.error.message}` : "").toBe(
        true
      );
      if (first.isErr()) return;
      const firstBytes = await fs.readFile(path.join(projectPath, first.value.appPackagePath));

      vi.setSystemTime(new Date("2027-01-01T00:00:00.000Z"));
      const second = await preparePublication({
        projectPath: secondProjectPath,
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
        tenantContext,
      });
      expect(second.isOk()).toBe(true);
      if (second.isErr()) return;
      const secondBytes = await fs.readFile(
        path.join(secondProjectPath, second.value.appPackagePath)
      );
      expect(secondBytes).toEqual(firstBytes);
    } finally {
      vi.useRealTimers();
    }
  });

  it("SPFX-PREP-05: rejects unqualified tenant-origin CSP fields", async () => {
    const projectPath = await fs.mkdtemp(
      path.join(os.tmpdir(), "prepared-publication-missing-env-")
    );
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(
      path.join(projectPath, "sharepoint", "solution.sppkg"),
      createSppkg({ includeCsp: true })
    );
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    process.env.TENANT_MCP_URL = "https://attacker.example.test/mcp";

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext,
    });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe("UnqualifiedTenantOriginField");
    expect(await fs.pathExists(path.join(projectPath, ".fx", "sharepoint-publications"))).toBe(
      false
    );
  });
});
