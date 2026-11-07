// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TeamsManifestConverter } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { createHash, randomBytes } from "crypto";
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

function incompressibleString(length: number): string {
  return randomBytes(Math.ceil((length * 3) / 4))
    .toString("base64")
    .slice(0, length);
}

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

async function createLegacyMaterializedApplicationPackage(
  kind: "agent" | "connector",
  sourcePlugin: Buffer
): Promise<{ bytes: Buffer; legacyId: string; legacyPath: string }> {
  const plugin = JSON.parse(sourcePlugin.toString("utf8"));
  plugin.runtimes[0].spec.url = `https://tenant.example.test/_api/mcp/spfx?solutionId=${encodeURIComponent(productId)}`;
  const materializedPlugin = Buffer.from(JSON.stringify(plugin));
  const contentHash = createHash("sha256").update(materializedPlugin).digest("hex").slice(0, 16);
  const legacyId = `spfx-api-${contentHash}`;
  const legacyPath = `plugins/${legacyId}.json`;
  const zip = new AdmZip(createApplicationPackage(kind));

  if (kind === "agent") {
    const agent = JSON.parse(zip.readAsText("declarativeAgent.json"));
    agent.actions.push({ id: legacyId, file: legacyPath });
    zip.updateFile("declarativeAgent.json", Buffer.from(JSON.stringify(agent)));
  } else {
    const manifest = JSON.parse(zip.readAsText("manifest.json"));
    manifest.agentConnectors.push({
      id: legacyId,
      displayName: "Legacy SharePoint plugin",
      toolSource: { plugin: { id: plugin.namespace, file: legacyPath } },
    });
    zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
  }
  zip.addFile(legacyPath, materializedPlugin);
  return { bytes: zip.toBuffer(), legacyId, legacyPath };
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

  function completeConnectorManifest(manifestVersion: string): Record<string, unknown> {
    const manifest = JSON.parse(
      new AdmZip(createApplicationPackage("connector")).readAsText("manifest.json")
    );
    delete manifest.agentSkills;
    return {
      ...manifest,
      manifestVersion,
      developer: {
        name: "Contoso",
        websiteUrl: "https://contoso.example.test",
        privacyUrl: "https://contoso.example.test/privacy",
        termsOfUseUrl: "https://contoso.example.test/terms",
      },
      description: { short: "Sample", full: "Sample application" },
      icons: { color: "color.png", outline: "outline.png" },
      accentColor: "#FFFFFF",
    };
  }

  function connectorApplication(manifest: Record<string, unknown>): Buffer {
    const zip = new AdmZip(createApplicationPackage("connector"));
    zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
    return zip.toBuffer();
  }

  it("SPFX-PREP-12: emits a connector representation accepted by the genuine Teams manifest converter", async () => {
    const result = await assembleSharePointApplicationPackage({
      applicationPackage: connectorApplication(completeConnectorManifest("devPreview")),
      sharePointPackage: createSppkg(),
    });

    expect(result.isOk()).toBe(true);
    if (result.isErr()) return;
    const manifestText = new AdmZip(result.value.packageBytes).readAsText("manifest.json");
    const typed = TeamsManifestConverter.jsonToManifest(manifestText);
    expect(typed.manifestVersion).toBe("devPreview");
    expect(JSON.parse(manifestText).agentConnectors[1].toolSource.plugin).toEqual({
      id: "sample-plugin",
      file: result.value.managedReference.path,
    });
  });

  it.each(["1.30", "1.28"])(
    "SPFX-PREP-12: rejects connector-only assembly for manifest version %s without plugin tool sources",
    async (manifestVersion) => {
      const manifest = { ...completeConnectorManifest(manifestVersion), agentConnectors: [] };
      expect(() => TeamsManifestConverter.jsonToManifest(JSON.stringify(manifest))).not.toThrow();
      expect(() =>
        TeamsManifestConverter.jsonToManifest(
          JSON.stringify({
            ...manifest,
            agentConnectors: [
              {
                id: "plugin-connector",
                displayName: "Plugin connector",
                toolSource: { plugin: { id: "sample-plugin", file: "plugins/sample.json" } },
              },
            ],
          })
        )
      ).toThrow(/toolSource\.plugin/);

      const result = await assembleSharePointApplicationPackage({
        applicationPackage: connectorApplication(manifest),
        sharePointPackage: createSppkg(),
      });

      expect(result.isErr()).toBe(true);
      if (result.isOk()) return;
      expect(result.error.code).toBe("UnsupportedAgentConnectorManifestVersion");
    }
  );

  it.each(["agent", "connector"] as const)(
    "SPFX-PREP-13: migrates a legacy materialized %s plugin without dangling references",
    async (kind) => {
      const sourceSharePoint = createSppkg();
      const inspected = await inspectSharePointPackage(sourceSharePoint);
      expect(inspected.isOk()).toBe(true);
      if (inspected.isErr()) return;
      const legacy = await createLegacyMaterializedApplicationPackage(
        kind,
        inspected.value.pluginManifest
      );
      const result = await assembleSharePointApplicationPackage({
        applicationPackage: legacy.bytes,
        sharePointPackage: sourceSharePoint,
      });

      expect(result.isOk()).toBe(true);
      if (result.isErr()) return;
      expect(legacy.legacyId).not.toBe(result.value.managedReference.id);
      const output = new AdmZip(result.value.packageBytes);
      const outputEntries = new Set(output.getEntries().map((entry) => entry.entryName));
      expect(outputEntries).not.toContain(legacy.legacyPath);
      expect(outputEntries).toContain(result.value.managedReference.path);
      expect(outputEntries).toContain(userPluginPath);
      expect(output.readAsText(userPluginPath)).toBe(JSON.stringify({ retained: true }));

      const manifest = JSON.parse(output.readAsText("manifest.json"));
      expect(manifest.agentSkills).toEqual([{ folder: "skills/existing" }]);
      const managedReferences =
        kind === "agent"
          ? JSON.parse(output.readAsText("declarativeAgent.json")).actions.filter(
              (action: { id?: string; file?: string }) =>
                action.id?.startsWith("spfx-api-") || action.file?.startsWith("plugins/spfx-api-")
            )
          : manifest.agentConnectors.filter(
              (connector: { id?: string; toolSource?: { plugin?: { file?: string } } }) =>
                connector.id?.startsWith("spfx-api-") ||
                connector.toolSource?.plugin?.file?.startsWith("plugins/spfx-api-")
            );
      expect(managedReferences).toEqual([
        kind === "agent"
          ? {
              id: result.value.managedReference.id,
              file: result.value.managedReference.path,
            }
          : expect.objectContaining({
              id: result.value.managedReference.id,
              toolSource: expect.objectContaining({
                plugin: expect.objectContaining({ file: result.value.managedReference.path }),
              }),
            }),
      ]);

      const referencedFiles =
        kind === "agent"
          ? JSON.parse(output.readAsText("declarativeAgent.json")).actions.map(
              (action: { file: string }) => action.file
            )
          : manifest.agentConnectors.flatMap(
              (connector: { toolSource?: { plugin?: { file?: string } } }) => {
                const file = connector.toolSource?.plugin?.file;
                return file ? [file] : [];
              }
            );
      expect(referencedFiles.every((file: string) => outputEntries.has(file))).toBe(true);
      if (kind === "agent") {
        expect(JSON.parse(output.readAsText("declarativeAgent.json")).instructions).toBe(
          "Keep these instructions"
        );
      } else {
        expect(manifest.agentConnectors[0].id).toBe("existing-connector");
      }
    }
  );

  it("SPFX-PREP-13: persists root-manifest cleanup when migrating a legacy connector to an agent action", async () => {
    const sourceSharePoint = createSppkg();
    const inspected = await inspectSharePointPackage(sourceSharePoint);
    expect(inspected.isOk()).toBe(true);
    if (inspected.isErr()) return;
    const legacy = await createLegacyMaterializedApplicationPackage(
      "connector",
      inspected.value.pluginManifest
    );
    const legacyZip = new AdmZip(legacy.bytes);
    const legacyManifest = JSON.parse(legacyZip.readAsText("manifest.json"));
    const legacyConnector = legacyManifest.agentConnectors.find(
      (connector: { id: string }) => connector.id === legacy.legacyId
    );
    const legacyPlugin = legacyZip.readFile(legacy.legacyPath);
    expect(legacyPlugin).not.toBeNull();
    if (legacyPlugin === null) return;
    const application = new AdmZip(createApplicationPackage("agent"));
    const manifest = JSON.parse(application.readAsText("manifest.json"));
    const unrelatedConnector = {
      id: "existing-connector",
      displayName: "Existing connector",
      toolSource: {
        remoteMcpServer: { mcpServerUrl: "https://existing.example.test/mcp" },
      },
    };
    manifest.agentConnectors = [unrelatedConnector, legacyConnector];
    application.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
    application.addFile(legacy.legacyPath, legacyPlugin);

    const result = await assembleSharePointApplicationPackage({
      applicationPackage: application.toBuffer(),
      sharePointPackage: sourceSharePoint,
    });

    expect(result.isOk()).toBe(true);
    if (result.isErr()) return;
    const output = new AdmZip(result.value.packageBytes);
    const outputManifest = JSON.parse(output.readAsText("manifest.json"));
    expect(outputManifest.agentConnectors).toEqual([unrelatedConnector]);
    expect(output.getEntries().map((entry) => entry.entryName)).not.toContain(legacy.legacyPath);
    const agent = JSON.parse(output.readAsText("declarativeAgent.json"));
    expect(agent.actions).toContainEqual({
      id: result.value.managedReference.id,
      file: result.value.managedReference.path,
    });
  });

  it.each(["application manifest", "declarative-agent manifest", "managed plugin"] as const)(
    "SPFX-PREP-14: rejects oversized %s metadata with a bounded error",
    async (surface) => {
      const sourceSharePoint = createSppkg();
      const inspected = await inspectSharePointPackage(sourceSharePoint);
      expect(inspected.isOk()).toBe(true);
      if (inspected.isErr()) return;
      const metadataLimit = SHAREPOINT_PACKAGE_LIMITS.metadataBytes;
      const application = new AdmZip(createApplicationPackage("agent"));

      if (surface === "application manifest") {
        const manifest = JSON.parse(application.readAsText("manifest.json"));
        manifest.description = incompressibleString(metadataLimit);
        application.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      } else if (surface === "declarative-agent manifest") {
        const agent = JSON.parse(application.readAsText("declarativeAgent.json"));
        agent.instructions = incompressibleString(metadataLimit);
        application.updateFile("declarativeAgent.json", Buffer.from(JSON.stringify(agent)));
      } else {
        const materializedPlugin = JSON.parse(inspected.value.pluginManifest.toString("utf8"));
        materializedPlugin.runtimes[0].spec.url = `https://tenant.example.test/_api/mcp/spfx?solutionId=${encodeURIComponent(productId)}`;
        materializedPlugin.metadataPadding = incompressibleString(metadataLimit);
        const pluginBytes = Buffer.from(JSON.stringify(materializedPlugin));
        const contentHash = createHash("sha256").update(pluginBytes).digest("hex").slice(0, 16);
        const legacyId = `spfx-api-${contentHash}`;
        const legacyPath = `plugins/${legacyId}.json`;
        const agent = JSON.parse(application.readAsText("declarativeAgent.json"));
        agent.actions.push({ id: legacyId, file: legacyPath });
        application.updateFile("declarativeAgent.json", Buffer.from(JSON.stringify(agent)));
        application.addFile(legacyPath, pluginBytes);
      }

      const result = await assembleSharePointApplicationPackage({
        applicationPackage: application.toBuffer(),
        sharePointPackage: sourceSharePoint,
      });

      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.code).toBe("MetadataTooLarge");
      }
    }
  );

  it("SPFX-PREP-14: rejects output metadata that exceeds the limit after adding the managed reference", async () => {
    const sourceSharePoint = createSppkg();
    const application = new AdmZip(createApplicationPackage("agent"));
    const agent = JSON.parse(application.readAsText("declarativeAgent.json"));
    const metadataLimit = SHAREPOINT_PACKAGE_LIMITS.metadataBytes;
    const emptyLength = Buffer.byteLength(JSON.stringify(agent));
    agent.instructions = incompressibleString(metadataLimit - emptyLength - 5);
    const sourceAgentBytes = Buffer.from(JSON.stringify(agent));
    expect(sourceAgentBytes.length).toBeLessThanOrEqual(metadataLimit);
    application.updateFile("declarativeAgent.json", sourceAgentBytes);

    const result = await assembleSharePointApplicationPackage({
      applicationPackage: application.toBuffer(),
      sharePointPackage: sourceSharePoint,
    });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.code).toBe("MetadataTooLarge");
    }
  });

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

  it("SPFX-PREP-14: rejects conflicting content at the deterministic managed plugin path", async () => {
    const sourceSharePoint = createSppkg();
    const inspected = await inspectSharePointPackage(sourceSharePoint);
    expect(inspected.isOk()).toBe(true);
    if (inspected.isErr()) return;
    const bindingHash = createHash("sha256").update(inspected.value.pluginManifest).digest("hex");
    const pluginId = `spfx-api-${bindingHash.slice(0, 16)}`;
    const pluginPath = `plugins/${pluginId}.json`;
    const application = new AdmZip(createApplicationPackage("agent"));
    const agent = JSON.parse(application.readAsText("declarativeAgent.json"));
    agent.actions.push({ id: pluginId, file: pluginPath });
    application.updateFile("declarativeAgent.json", Buffer.from(JSON.stringify(agent)));
    application.addFile(pluginPath, Buffer.from(JSON.stringify({ conflicting: true })));

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
