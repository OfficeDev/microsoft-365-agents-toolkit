// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { createHash } from "crypto";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  assembleSharePointApplicationPackage,
  inspectSharePointPackage,
  readSharePointPackage,
} from "../../../../../src/sharepointPackage";
import {
  readBoundedArchive,
  SHAREPOINT_PACKAGE_LIMITS,
} from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";

const productId = "11111111-2222-4333-8444-555555555555";
const appId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";
const userPluginPath = "plugins/user-plugin.json";
const agentPackageName = "ClientSideAssets/copilot-agent.zip";
const componentManifestName =
  "bbbbbbbb-cccc-4ddd-8eee-ffffffffffff/CopilotComponent_bbbbbbbb-cccc-4ddd-8eee-ffffffffffff.xml";

function createAgentPackage(): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    "ai-plugin.json",
    Buffer.from(
      JSON.stringify({
        schema_version: "v2.1",
        namespace: "sample-plugin",
        name_for_human: "Component actions",
        runtimes: [
          {
            type: "RemoteMCPServer",
            spec: { url: "{{TENANT_MCP_URL}}" },
          },
        ],
      })
    )
  );
  zip.addFile("manifest.json", Buffer.from(JSON.stringify({ id: "agent-only-id" })));
  zip.addFile(
    "declarativeAgent.json",
    Buffer.from(JSON.stringify({ instructions: "Do not import" }))
  );
  return zip.toBuffer();
}

function createSppkg(options: { agentPackages?: string[] } = {}): Buffer {
  const nested = createAgentPackage();
  const zip = new AdmZip();
  zip.addFile(
    "AppManifest.xml",
    Buffer.from(
      `<App ProductID="{${productId}}" Version="1.2.3.4" IsClientSideSolution="true" SkipFeatureDeployment="false" />`
    )
  );
  zip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="zip" ContentType="application/octet-stream" /></Types>`
    )
  );
  zip.addFile(
    "_rels/.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="package-manifest" Target="/AppManifest.xml" Id="r0" /></Relationships>`
    )
  );
  zip.addFile(
    "_rels/AppManifest.xml.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="manifest-clientsideasset" Target="/ClientSideAssets.xml" Id="r1" /></Relationships>`
    )
  );
  const packages = options.agentPackages ?? [agentPackageName];
  const relationships = packages
    .map(
      (packageName, index) =>
        `<Relationship Type="clientsideasset" Target="/${packageName}" Id="r${index + 2}" />`
    )
    .join("");
  zip.addFile(
    "ClientSideAssets/_rels/ClientSideAssets.xml.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}<Relationship Type="clientsideasset" Target="/ClientSideAssets/component.js" Id="script" /></Relationships>`
    )
  );
  zip.addFile("ClientSideAssets.xml", Buffer.from("<ClientSideAssets />"));
  zip.addFile(componentManifestName, Buffer.from("<Elements><ClientSideComponent /></Elements>"));
  zip.addFile("ClientSideAssets/component.js", Buffer.from("component asset"));
  for (const packageName of packages) {
    zip.addFile(packageName, nested);
  }
  return zip.toBuffer();
}

function createApplicationPackage(kind: "agent" | "connector"): Buffer {
  const zip = new AdmZip();
  const manifest: Record<string, unknown> = {
    manifestVersion: kind === "agent" ? "1.30" : "devPreview",
    version: "2.3.4",
    id: appId,
    name: { short: "Sample", full: "Sample application" },
    agentSkills: [{ folder: "skills/existing" }],
    ...(kind === "agent"
      ? {
          copilotAgents: {
            declarativeAgents: [{ id: "agent", file: "declarativeAgent.json" }],
          },
        }
      : {
          agentConnectors: [
            {
              id: "existing-connector",
              displayName: "Existing connector",
              toolSource: {
                remoteMcpServer: { mcpServerUrl: "https://existing.example.test/mcp" },
              },
            },
          ],
        }),
  };
  zip.addFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
  if (kind === "agent") {
    zip.addFile(
      "declarativeAgent.json",
      Buffer.from(
        JSON.stringify({
          version: "v1.9",
          name: "Existing agent",
          description: "Existing agent description",
          instructions: "Keep these instructions",
          actions: [{ id: "existing-action", file: userPluginPath }],
        })
      )
    );
  }
  zip.addFile(userPluginPath, Buffer.from(JSON.stringify({ retained: true })));
  return zip.toBuffer();
}

describe("SharePoint package inspection", () => {
  const directories: string[] = [];

  afterEach(async () => {
    await Promise.all(directories.splice(0).map((directory) => fs.remove(directory)));
  });

  it("SPFX-PREP-01: inspects the exact bounded package and extracts only its API plugin", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-package-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    const original = createSppkg();
    await fs.writeFile(sourcePath, original);

    const result = await readSharePointPackage(sourcePath);

    expect(result.isOk()).toBe(true);
    if (result.isErr()) return;
    expect(result.value.productId).toBe(productId);
    expect(result.value.version).toBe("1.2.3.4");
    expect(result.value.skipFeatureDeployment).toBe(false);
    expect(result.value.sourceSha256).toBe(createHash("sha256").update(original).digest("hex"));
    expect(result.value.pluginManifest.toString("utf8")).toContain("TENANT_MCP_URL");
    expect(result.value.deploymentBytes).not.toEqual(original);
    expect(await fs.readFile(sourcePath)).toEqual(original);
  });

  it("SPFX-PREP-03: removes the embedded package and repairs OPC relationships", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-opc-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    await fs.writeFile(sourcePath, createSppkg());

    const result = await readSharePointPackage(sourcePath);

    expect(result.isOk()).toBe(true);
    if (result.isErr()) return;
    const deployment = new AdmZip(result.value.deploymentBytes);
    const names = deployment.getEntries().map((entry) => entry.entryName);
    expect(names).not.toContain(agentPackageName);
    expect(names).toContain("ClientSideAssets/component.js");
    expect(
      deployment.readAsText("ClientSideAssets/_rels/ClientSideAssets.xml.rels").toLowerCase()
    ).not.toContain("copilot-agent.zip");
    expect(deployment.readAsText("[Content_Types].xml")).not.toContain('Extension="zip"');
  });

  it.each(["agent", "connector"] as const)(
    "SPFX-PREP-12: assembles a tenant-neutral package while preserving %s content offline",
    async (kind) => {
      const sourceSharePoint = createSppkg();
      const sourceApplication = createApplicationPackage(kind);
      const originalSharePoint = Buffer.from(sourceSharePoint);
      const originalApplication = Buffer.from(sourceApplication);
      const openSpy = vi.spyOn(fs, "open");
      const readSpy = vi.spyOn(fs, "read");
      const writeSpy = vi.spyOn(fs, "writeFile");
      const fetchSpy = vi.spyOn(globalThis, "fetch");

      try {
        const result = await assembleSharePointApplicationPackage({
          applicationPackage: sourceApplication,
          sharePointPackage: sourceSharePoint,
        });

        expect(result.isOk()).toBe(true);
        if (result.isErr()) return;
        expect(result.value.sourceSharePointSha256).toBe(
          createHash("sha256").update(sourceSharePoint).digest("hex")
        );
        expect(result.value.sourceApplicationSha256).toBe(
          createHash("sha256").update(sourceApplication).digest("hex")
        );
        expect(result.value.packageSha256).toBe(
          createHash("sha256").update(result.value.packageBytes).digest("hex")
        );
        expect(result.value.productId).toBe(productId);
        expect(result.value.appId).toBe(appId);
        expect(result.value.managedReference.kind).toBe(
          kind === "agent" ? "declarativeAgentAction" : "agentConnector"
        );
        const output = new AdmZip(result.value.packageBytes);
        const manifest = JSON.parse(output.readAsText("manifest.json"));
        expect(manifest.agentSkills).toEqual([{ folder: "skills/existing" }]);
        expect(output.readAsText(userPluginPath)).toBe(JSON.stringify({ retained: true }));
        const plugin = JSON.parse(output.readAsText(result.value.managedReference.path));
        expect(plugin.runtimes[0].spec.url).toBe("{{TENANT_MCP_URL}}");
        if (kind === "agent") {
          const agent = JSON.parse(output.readAsText("declarativeAgent.json"));
          expect(agent.instructions).toBe("Keep these instructions");
          expect(agent.actions).toEqual([
            { id: "existing-action", file: userPluginPath },
            { id: result.value.managedReference.id, file: result.value.managedReference.path },
          ]);
        } else {
          expect(manifest.agentConnectors).toHaveLength(2);
          expect(manifest.agentConnectors[0].id).toBe("existing-connector");
          expect(manifest.agentConnectors[1].id).toBe(result.value.managedReference.id);
        }
        const repeated = await assembleSharePointApplicationPackage({
          applicationPackage: result.value.packageBytes,
          sharePointPackage: sourceSharePoint,
        });
        expect(repeated.isOk()).toBe(true);
        if (repeated.isOk()) {
          expect(repeated.value.packageBytes).toEqual(result.value.packageBytes);
        }
        expect(sourceSharePoint).toEqual(originalSharePoint);
        expect(sourceApplication).toEqual(originalApplication);
        expect(openSpy).not.toHaveBeenCalled();
        expect(readSpy).not.toHaveBeenCalled();
        expect(writeSpy).not.toHaveBeenCalled();
        expect(fetchSpy).not.toHaveBeenCalled();
      } finally {
        openSpy.mockRestore();
        readSpy.mockRestore();
        writeSpy.mockRestore();
        fetchSpy.mockRestore();
      }
    }
  );

  it("SPFX-PREP-14: rejects malformed and oversized application inputs", async () => {
    const malformed = await assembleSharePointApplicationPackage({
      applicationPackage: Buffer.from("not a ZIP package"),
      sharePointPackage: createSppkg(),
    });
    expect(malformed.isErr()).toBe(true);
    if (malformed.isErr()) {
      expect(malformed.error.code).toBe("InvalidArchive");
    }

    const originalLimit = SHAREPOINT_PACKAGE_LIMITS.inputBytes;
    SHAREPOINT_PACKAGE_LIMITS.inputBytes = 1024;
    try {
      const oversized = await assembleSharePointApplicationPackage({
        applicationPackage: createApplicationPackage("agent"),
        sharePointPackage: createSppkg(),
      });
      expect(oversized.isErr()).toBe(true);
      if (oversized.isErr()) {
        expect(oversized.error.code).toBe("ArchiveTooLarge");
      }
    } finally {
      SHAREPOINT_PACKAGE_LIMITS.inputBytes = originalLimit;
    }
  });

  it("SPFX-PREP-14: rejects a conflicting managed plugin reference", async () => {
    const sourceSharePoint = createSppkg();
    const inspected = await inspectSharePointPackage(sourceSharePoint);
    expect(inspected.isOk()).toBe(true);
    if (inspected.isErr()) return;
    const pluginHash = createHash("sha256").update(inspected.value.pluginManifest).digest("hex");
    const pluginId = `spfx-api-${pluginHash.slice(0, 16)}`;
    const application = new AdmZip(createApplicationPackage("agent"));
    const agent = JSON.parse(application.readAsText("declarativeAgent.json"));
    agent.actions.push({ id: pluginId, file: "plugins/conflicting.json" });
    application.updateFile("declarativeAgent.json", Buffer.from(JSON.stringify(agent)));

    const result = await assembleSharePointApplicationPackage({
      applicationPackage: application.toBuffer(),
      sharePointPackage: sourceSharePoint,
    });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.code).toBe("ApplicationPackageCollision");
    }
  });

  it("SPFX-PREP-04: rejects ambiguous nested agent packages instead of selecting one", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-ambiguous-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    await fs.writeFile(
      sourcePath,
      createSppkg({
        agentPackages: ["ClientSideAssets/copilot-agent.zip", "ClientSideAssets/second-agent.zip"],
      })
    );

    const result = await readSharePointPackage(sourcePath);

    expect(result.isErr()).toBe(true);
  });

  it("SPFX-PREP-01: rejects an SPPKG without a component manifest", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-no-component-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    const zip = new AdmZip(createSppkg());
    zip.deleteFile(componentManifestName);
    await fs.writeFile(sourcePath, zip.toBuffer());

    const result = await readSharePointPackage(sourcePath);

    expect(result.isErr()).toBe(true);
  });

  it("SPFX-PREP-04: rejects malformed package archives without extracting entries", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-invalid-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    await fs.writeFile(sourcePath, Buffer.from("not a ZIP package"));

    const result = await readSharePointPackage(sourcePath);

    expect(result.isErr()).toBe(true);
  });

  it("SPFX-PREP-04: rejects a package that changes during its bounded read", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-package-race-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    await fs.writeFile(sourcePath, createSppkg());
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
      const result = await readSharePointPackage(sourcePath);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.code).toBe("PackageChangedDuringRead");
      }
    } finally {
      readSpy.mockRestore();
    }
  });

  it("SPFX-PREP-04: caps bytes read when a package grows after its initial size check", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-package-growth-"));
    directories.push(directory);
    const sourcePath = path.join(directory, "solution.sppkg");
    const original = createSppkg();
    const originalLimit = SHAREPOINT_PACKAGE_LIMITS.inputBytes;
    const testLimit = 16 * 1024;
    expect(original.length).toBeLessThan(testLimit);
    SHAREPOINT_PACKAGE_LIMITS.inputBytes = testLimit;
    await fs.writeFile(sourcePath, original);
    const originalRead = fs.read.bind(fs);
    let grew = false;
    let bytesRead = 0;
    const readSpy = vi
      .spyOn(fs, "read")
      .mockImplementation(async (fd, buffer, offset, length, position) => {
        if (!grew) {
          grew = true;
          await fs.truncate(sourcePath, testLimit + 100);
        }
        const result = await originalRead(fd, buffer, offset, length, position);
        bytesRead += result.bytesRead;
        return result;
      });

    try {
      const result = await readSharePointPackage(sourcePath);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.code).toBe("ArchiveTooLarge");
      }
      expect(bytesRead).toBe(testLimit + 1);
      expect(readSpy).toHaveBeenCalled();
    } finally {
      readSpy.mockRestore();
      SHAREPOINT_PACKAGE_LIMITS.inputBytes = originalLimit;
    }
  });

  it("SPFX-PREP-04: rejects traversal, duplicate paths, and oversized archives", () => {
    const traversal = new AdmZip();
    traversal.addFile("outside.txt", Buffer.from("outside"));
    const traversalEntry = traversal.getEntry("outside.txt");
    expect(traversalEntry).toBeDefined();
    if (traversalEntry) traversalEntry.entryName = "../outside.txt";
    const traversalResult = readBoundedArchive(traversal.toBuffer());
    expect(traversalResult.isErr()).toBe(true);

    const duplicate = new AdmZip();
    duplicate.addFile("content.xml", Buffer.from("first"));
    duplicate.addFile("CONTENT.XML", Buffer.from("second"));
    const duplicateResult = readBoundedArchive(duplicate.toBuffer());
    expect(duplicateResult.isErr()).toBe(true);

    const oversizedResult = readBoundedArchive(
      Buffer.alloc(SHAREPOINT_PACKAGE_LIMITS.inputBytes + 1)
    );
    expect(oversizedResult.isErr()).toBe(true);
  });
});
