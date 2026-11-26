// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { randomBytes } from "crypto";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  inspectSharePointPackage,
  inspectSharePointPackageIdentity,
  readBoundedArchive,
  readSharePointPackage,
  SHAREPOINT_PACKAGE_LIMITS,
} from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";
import { componentId, createSppkg, productId } from "./preparedPublicationFixtures";

const componentPath = `${componentId}/CopilotComponent_${componentId}.xml`;
const initialLimits = { ...SHAREPOINT_PACKAGE_LIMITS };

function archive(...entries: [string, Buffer][]): Buffer {
  const zip = new AdmZip();
  for (const [name, bytes] of entries) zip.addFile(name, bytes);
  return zip.toBuffer();
}

describe("Bounded SharePoint archives and OPC metadata", () => {
  afterEach(() => {
    Object.assign(SHAREPOINT_PACKAGE_LIMITS, initialLimits);
    vi.restoreAllMocks();
  });

  it.each(["/absolute.xml", "C:/drive.xml", "bad\\path.xml", "./part.xml", "a/../part.xml"])(
    "SPFX-PREP-04: rejects noncanonical archive path %s",
    (name) => {
      const zip = new AdmZip();
      zip.addFile("part.xml", Buffer.from("data"));
      const entry = zip.getEntry("part.xml");
      if (!entry) throw new Error("Missing fixture entry.");
      entry.entryName = name;
      const result = readBoundedArchive(zip.toBuffer());
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.code).toBe("InvalidArchive");
    }
  );

  it("SPFX-PREP-04: ignores canonical directory entries without treating them as content", () => {
    const result = readBoundedArchive(
      archive(["assets/", Buffer.alloc(0)], ["assets/part.xml", Buffer.from("data")])
    );
    expect(result.isErr() ? result.error : result.value).toEqual([
      { name: "assets/part.xml", data: Buffer.from("data") },
    ]);
  });

  it.each([
    { limit: "entries", value: 1, code: "TooManyEntries" },
    { limit: "entryBytes", value: 15, code: "ArchiveEntryTooLarge" },
    { limit: "expandedBytes", value: 31, code: "ArchiveExpansionTooLarge" },
    { limit: "compressionRatio", value: 0.1, code: "ArchiveEntryTooLarge" },
  ])(
    "SPFX-PREP-04: enforces the $limit boundary before extracting content",
    ({ limit, value, code }) => {
      const bytes = archive(["first", randomBytes(16)], ["second", randomBytes(16)]);
      Object.assign(SHAREPOINT_PACKAGE_LIMITS, { [limit]: value });
      const result = readBoundedArchive(bytes);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.code).toBe(code);
    }
  );

  it.each([
    { method: 12, size: 4, compressedSize: undefined, code: "UnsupportedCompression" },
    { method: 8, size: 0, compressedSize: undefined, code: "ArchiveEntryTooLarge" },
    { method: 0, size: 4, compressedSize: 0, code: "ArchiveEntryTooLarge" },
  ])(
    "SPFX-PREP-04: refuses unsafe ZIP headers ($code, method $method)",
    ({ method, size, compressedSize, code }) => {
      const bytes = archive(["part", Buffer.from("data")]);
      const central = bytes.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
      expect(central).toBeGreaterThan(0);
      bytes.writeUInt16LE(method, central + 10);
      bytes.writeUInt32LE(size, central + 24);
      if (compressedSize !== undefined) bytes.writeUInt32LE(compressedSize, central + 20);
      const result = readBoundedArchive(bytes);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.code).toBe(code);
    }
  );

  it("SPFX-PREP-04: rejects a corrupt central directory independently of the ZIP envelope", () => {
    const bytes = archive(["part", Buffer.from("data")]);
    const central = bytes.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
    bytes.writeUInt32LE(0, central);
    const result = readBoundedArchive(bytes);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe("InvalidArchive");
  });

  it("SPFX-PREP-04: rejects corrupt entry content rather than returning unverified bytes", () => {
    const bytes = archive(["part", randomBytes(64)]);
    const nameLength = bytes.readUInt16LE(26);
    const extraLength = bytes.readUInt16LE(28);
    const content = 30 + nameLength + extraLength;
    bytes[content] ^= 0xff;
    const result = readBoundedArchive(bytes);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe("InvalidArchive");
  });

  it.each([
    {
      name: "missing application manifest",
      part: "AppManifest.xml",
      data: undefined,
      code: "MissingAppManifest",
    },
    {
      name: "malformed application XML",
      part: "AppManifest.xml",
      data: "<App",
      code: "InvalidPackageMetadata",
    },
    {
      name: "missing product attributes",
      part: "AppManifest.xml",
      data: "<App />",
      code: "InvalidAppManifest",
    },
    {
      name: "foreign application root",
      part: "AppManifest.xml",
      data: "<Foreign />",
      code: "InvalidAppManifest",
    },
    {
      name: "invalid solution version",
      part: "AppManifest.xml",
      data: `<App ProductID="${productId}" Version="1.2.3" SkipFeatureDeployment="false" />`,
      code: "InvalidAppManifest",
    },
    {
      name: "invalid deployment flag",
      part: "AppManifest.xml",
      data: `<App ProductID="${productId}" Version="1.2.3.4" SkipFeatureDeployment="sometimes" />`,
      code: "InvalidAppManifest",
    },
    {
      name: "malformed component XML",
      part: componentPath,
      data: "<Elements",
      code: "InvalidPackageMetadata",
    },
    {
      name: "missing component element",
      part: componentPath,
      data: "<Elements><Foreign /></Elements>",
      code: "InvalidCopilotComponent",
    },
    {
      name: "missing content types",
      part: "[Content_Types].xml",
      data: undefined,
      code: "MissingContentTypes",
    },
    {
      name: "malformed content types",
      part: "[Content_Types].xml",
      data: "<Types",
      code: "InvalidPackageMetadata",
    },
    {
      name: "foreign content type root",
      part: "[Content_Types].xml",
      data: "<Foreign />",
      code: "InvalidContentTypes",
    },
    {
      name: "malformed discovery relationship",
      part: "_rels/.rels",
      data: "<Relationships",
      code: "InvalidPackageMetadata",
    },
    {
      name: "foreign discovery relationship root",
      part: "_rels/.rels",
      data: "<Foreign />",
      code: "InvalidPackageMetadata",
    },
    {
      name: "malformed later relationship",
      part: "zz/_rels/late.xml.rels",
      data: "<Relationships",
      code: "InvalidPackageMetadata",
    },
    {
      name: "foreign later relationship root",
      part: "zz/_rels/late.xml.rels",
      data: "<Foreign />",
      code: "InvalidPackageMetadata",
    },
    {
      name: "scalar later relationship root",
      part: "zz/_rels/late.xml.rels",
      data: "<Relationships>scalar</Relationships>",
      code: "InvalidPackageMetadata",
    },
  ])(
    "SPFX-PREP-04: rejects $name before producing a deployment copy",
    async ({ part, data, code }) => {
      const source = new AdmZip(createSppkg());
      if (data === undefined) source.deleteFile(part);
      else source.addFile(part, Buffer.from(data));
      const bytes = source.toBuffer();
      const original = Buffer.from(bytes);
      const result = await inspectSharePointPackage(bytes);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.code).toBe(code);
      expect(bytes).toEqual(original);
    }
  );

  it("SPFX-PREP-04: enforces the XML metadata cap before parsing a component", async () => {
    const zip = new AdmZip(createSppkg());
    zip.updateFile(componentPath, randomBytes(SHAREPOINT_PACKAGE_LIMITS.metadataBytes + 1));
    const result = await inspectSharePointPackage(zip.toBuffer());
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe("MetadataTooLarge");
  });

  it.each([
    { name: "malformed nested archive", plugin: undefined, code: "InvalidArchive" },
    { name: "missing plugin", plugin: "", code: "EmbeddedPluginNotUnique" },
    { name: "malformed plugin JSON", plugin: "{", code: "InvalidPluginManifest" },
    { name: "non-object plugin", plugin: "null", code: "UnsupportedPluginRuntime" },
    {
      name: "unsupported runtime",
      plugin: '{"runtimes":[{"type":"OpenApi"}]}',
      code: "UnsupportedPluginRuntime",
    },
  ])("SPFX-PREP-04: rejects $name", async ({ plugin, code }) => {
    const zip = new AdmZip(createSppkg());
    const nested = new AdmZip();
    if (plugin) nested.addFile("ai-plugin.json", Buffer.from(plugin));
    zip.updateFile(
      "ClientSideAssets/agent.zip",
      plugin === undefined ? Buffer.from("not ZIP") : nested.toBuffer()
    );
    const result = await inspectSharePointPackage(zip.toBuffer());
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe(code);
  });

  it("SPFX-PREP-04: enforces the nested plugin metadata cap independently of archive expansion", async () => {
    const zip = new AdmZip(createSppkg());
    const nested = new AdmZip();
    nested.addFile("ai-plugin.json", randomBytes(SHAREPOINT_PACKAGE_LIMITS.metadataBytes + 1));
    zip.updateFile("ClientSideAssets/agent.zip", nested.toBuffer());
    const result = await inspectSharePointPackage(zip.toBuffer());
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe("MetadataTooLarge");
  });

  it.each([
    "",
    "\\invalid",
    "%zz",
    "//outside.test/agent.zip",
    "../../outside.zip",
    "https://outside.test/agent.zip",
  ])("SPFX-PREP-04: does not accept an invalid or foreign OPC target (%s)", async (target) => {
    const zip = new AdmZip(createSppkg());
    zip.updateFile(
      "_rels/ClientSideAssets.xml.rels",
      Buffer.from(`<Relationships><Relationship Target="${target}" /></Relationships>`)
    );
    const result = await inspectSharePointPackage(zip.toBuffer());
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.code).toBe("PluginPackageNotLinked");
  });

  it("SPFX-PREP-03: repairs relative and escaped OPC targets, preserving external and unrelated metadata", async () => {
    const zip = new AdmZip(createSppkg());
    const rels = `<Relationships><Relationship Id="agent" Target="ClientSideAssets/%61gent.zip" /><Relationship Id="external" Target="/ClientSideAssets/agent.zip" TargetMode="External" /><Relationship Id="component" Target="ClientSideAssets/component.js" /></Relationships>`;
    zip.updateFile("_rels/ClientSideAssets.xml.rels", Buffer.from(rels));
    zip.addFile("not-an-opc-location.rels", Buffer.from("must not be interpreted as XML"));
    zip.addFile("zz/_rels/empty.xml.rels", Buffer.from("<Relationships><Other /></Relationships>"));
    zip.updateFile(
      "[Content_Types].xml",
      Buffer.from(
        `<Types><Default Extension="zip" ContentType="application/zip" /><Override PartName="/ClientSideAssets/agent.zip" ContentType="application/zip" /><Override PartName="/ClientSideAssets/component.js" ContentType="text/javascript" /></Types>`
      )
    );
    const otherZip = archive(["unrelated.txt", Buffer.from("unrelated")]);
    zip.addFile("ClientSideAssets/other.zip", otherZip);

    const result = await inspectSharePointPackage(zip.toBuffer());
    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    if (result.isErr()) return;
    const output = new AdmZip(result.value.deploymentBytes);
    expect(output.getEntry("ClientSideAssets/agent.zip")).toBeNull();
    expect(output.readAsText("_rels/ClientSideAssets.xml.rels")).not.toContain('Id="agent"');
    expect(output.readAsText("_rels/ClientSideAssets.xml.rels")).toContain('Id="external"');
    expect(output.readAsText("[Content_Types].xml")).not.toContain(
      'PartName="/ClientSideAssets/agent.zip"'
    );
    expect(output.readAsText("[Content_Types].xml")).toContain(
      'PartName="/ClientSideAssets/component.js"'
    );
    expect(output.readAsText("[Content_Types].xml")).toContain('Extension="zip"');
    expect(output.readFile("ClientSideAssets/other.zip")).toEqual(otherZip);
    expect(output.readAsText("not-an-opc-location.rels")).toBe("must not be interpreted as XML");
  });

  it("SPFX-PREP-01: supports a package with tenant-wide deployment enabled", async () => {
    const zip = new AdmZip(createSppkg());
    zip.updateFile(
      "AppManifest.xml",
      Buffer.from(`<App ProductID="${productId}" Version="1.2.3.4" SkipFeatureDeployment="true" />`)
    );
    const result = await inspectSharePointPackageIdentity(zip.toBuffer());
    expect(result.isErr() ? result.error : result.value).toEqual({
      productId,
      version: "1.2.3.4",
      skipFeatureDeployment: true,
    });
    expect((await inspectSharePointPackageIdentity(Buffer.from("invalid"))).isErr()).toBe(true);
  });

  it("SPFX-PREP-04: closes the descriptor when the initial file size already exceeds the limit", async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), "spfx-bounds-"));
    try {
      const file = path.join(directory, "solution.sppkg");
      await fs.writeFile(file, createSppkg());
      SHAREPOINT_PACKAGE_LIMITS.inputBytes = 1;
      const close = vi.spyOn(fs, "close");
      const result = await readSharePointPackage(file);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.code).toBe("ArchiveTooLarge");
      expect(close).toHaveBeenCalledOnce();
      const missing = await readSharePointPackage(path.join(directory, "missing.sppkg"));
      expect(missing.isErr()).toBe(true);
      if (missing.isErr()) expect(missing.error.code).toBe("PackageReadFailed");
    } finally {
      await fs.remove(directory);
    }
  });
});
