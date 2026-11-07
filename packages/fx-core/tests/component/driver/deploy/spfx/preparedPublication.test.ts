// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it } from "vitest";
import {
  compareSharePointPackageVersions,
  isAmbiguousPublicationState,
  loadPreparedPublication,
  preparePublication,
  updatePreparedPublication,
} from "../../../../../src/component/driver/deploy/spfx/utility/preparedPublication";

const appId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";
const productId = "11111111-2222-4333-8444-555555555555";

function createSppkg(): Buffer {
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
  packageZip.addFile("ClientSideAssets/agent.zip", agent.toBuffer());
  packageZip.addFile("ClientSideAssets/component.js", Buffer.from("component"));
  return packageZip.toBuffer();
}

function createAppPackage(): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    "manifest.json",
    Buffer.from(
      JSON.stringify({
        manifestVersion: "1.30",
        version: "2.3.4",
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
  zip.addFile("instruction.txt", Buffer.from("Original application instructions"));
  zip.addFile("color.png", Buffer.from("icon"));
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

  it("SPFX-PREP-02/05: freezes and journals the derived artifacts without changing sources", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-publication-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    const sharePointSource = createSppkg();
    const appSource = createAppPackage();
    const sharePointPath = path.join(projectPath, "sharepoint", "solution.sppkg");
    const appPath = path.join(projectPath, "appPackage", "application.zip");
    await fs.writeFile(sharePointPath, sharePointSource);
    await fs.writeFile(appPath, appSource);
    process.env.TENANT_MCP_URL = "https://tenant.example.test/mcp";

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
    });

    expect(result.isOk()).toBe(true);
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
    expect(JSON.stringify(loaded.value.journal)).not.toContain("tenant.example.test");

    const applicationPath = path.join(projectPath, result.value.appPackagePath);
    const stagedApp = new AdmZip(await fs.readFile(applicationPath));
    const stagedAgent = JSON.parse(stagedApp.readAsText("declarativeAgent.json"));
    expect(stagedAgent.instructions).toBe("Keep these instructions");
    expect(stagedAgent.actions).toHaveLength(1);
    const pluginPath = stagedAgent.actions[0].file;
    expect(stagedApp.readAsText(pluginPath)).toContain("https://tenant.example.test/mcp");

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
    });
    expect(retry.isErr()).toBe(true);
    if (retry.isErr()) {
      expect(retry.error.code).toBe("UnresolvedPreparedPublication");
    }
  });

  it("SPFX-PREP-05: rejects an unresolved tenant value without staging outputs", async () => {
    const projectPath = await fs.mkdtemp(
      path.join(os.tmpdir(), "prepared-publication-missing-env-")
    );
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    delete process.env.TENANT_MCP_URL;

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
    });

    expect(result.isErr()).toBe(true);
    expect(await fs.pathExists(path.join(projectPath, ".fx", "sharepoint-publications"))).toBe(
      false
    );
  });
});
