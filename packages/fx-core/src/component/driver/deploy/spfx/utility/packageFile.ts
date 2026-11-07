// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { createHash } from "crypto";
import fs from "fs-extra";
import path from "path";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";
import xml2js from "xml2js";

export const SHAREPOINT_PACKAGE_LIMITS = {
  inputBytes: 100 * 1024 * 1024,
  entries: 10_000,
  expandedBytes: 512 * 1024 * 1024,
  entryBytes: 100 * 1024 * 1024,
  compressionRatio: 100,
  metadataBytes: 1024 * 1024,
};

export class SharePointPackageInspectionError extends Error {
  public constructor(
    public readonly code: string,
    message: string
  ) {
    super(message);
    this.name = "SharePointPackageInspectionError";
  }
}

export interface SharePointPackageInspection {
  sourceSha256: string;
  deploymentBytes: Buffer;
  deploymentSha256: string;
  productId: string;
  version: string;
  skipFeatureDeployment: boolean;
  embeddedPackagePath: string;
  pluginManifest: Buffer;
}

export interface SharePointPackageIdentity {
  productId: string;
  version: string;
  skipFeatureDeployment: boolean;
}

export interface ArchiveEntry {
  name: string;
  data: Buffer;
}

function packageError(code: string, message: string): SharePointPackageInspectionError {
  return new SharePointPackageInspectionError(code, message);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isRecordArray(value: unknown): value is Record<string, unknown>[] {
  return Array.isArray(value) && value.every(isRecord);
}

function getAttribute(record: Record<string, unknown>, name: string): string | undefined {
  const attributes = record["$"];
  if (!isRecord(attributes)) {
    return undefined;
  }
  const value = attributes[name];
  return typeof value === "string" ? value : undefined;
}

function sha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function normalizeEntryName(name: string, isDirectory: boolean): string {
  const normalizedSeparators = name;
  if (
    normalizedSeparators.includes("\\") ||
    normalizedSeparators.includes("\0") ||
    normalizedSeparators.startsWith("/") ||
    /^[A-Za-z]:/.test(normalizedSeparators)
  ) {
    throw packageError("InvalidArchivePath", "The package contains an invalid entry path.");
  }
  const entryPath =
    isDirectory && normalizedSeparators.endsWith("/")
      ? normalizedSeparators.slice(0, -1)
      : normalizedSeparators;
  const normalized = path.posix.normalize(entryPath);
  if (normalized === ".." || normalized.startsWith("../") || normalized === ".") {
    throw packageError("InvalidArchivePath", "The package contains a path outside its root.");
  }
  if (normalized !== entryPath) {
    throw packageError("InvalidArchivePath", "The package contains a non-canonical entry path.");
  }
  return normalized;
}

export function readBoundedArchive(
  bytes: Buffer
): Result<ArchiveEntry[], SharePointPackageInspectionError> {
  if (bytes.length > SHAREPOINT_PACKAGE_LIMITS.inputBytes) {
    return err(packageError("ArchiveTooLarge", "The package exceeds the supported input size."));
  }

  let zip: AdmZip;
  try {
    zip = new AdmZip(bytes);
  } catch {
    return err(packageError("InvalidArchive", "The package is not a valid ZIP archive."));
  }

  let zipEntries: ReturnType<AdmZip["getEntries"]>;
  try {
    zipEntries = zip.getEntries();
  } catch {
    return err(packageError("InvalidArchive", "The package is not a valid ZIP archive."));
  }
  if (zipEntries.length > SHAREPOINT_PACKAGE_LIMITS.entries) {
    return err(packageError("TooManyEntries", "The package contains too many entries."));
  }
  let expandedBytes = 0;
  const seen = new Set<string>();
  const entries: ArchiveEntry[] = [];
  try {
    for (const entry of zipEntries) {
      const name = normalizeEntryName(entry.entryName, entry.isDirectory);
      const duplicateKey = name.toLocaleLowerCase("en-US");
      if (seen.has(duplicateKey)) {
        return err(packageError("DuplicateArchivePath", "The package contains duplicate paths."));
      }
      seen.add(duplicateKey);
      if (entry.isDirectory) {
        continue;
      }
      const compressedSize = entry.header.compressedSize;
      const uncompressedSize = entry.header.size;
      const method = entry.header.method;
      if (method !== 0 && method !== 8) {
        return err(
          packageError("UnsupportedCompression", "The package uses unsupported compression.")
        );
      }
      if (
        (method === 8 && compressedSize > 0 && uncompressedSize === 0) ||
        uncompressedSize > SHAREPOINT_PACKAGE_LIMITS.entryBytes ||
        (compressedSize === 0 && uncompressedSize > 0) ||
        (compressedSize > 0 &&
          uncompressedSize / compressedSize > SHAREPOINT_PACKAGE_LIMITS.compressionRatio)
      ) {
        return err(packageError("ArchiveEntryTooLarge", "A package entry exceeds safe limits."));
      }
      expandedBytes += uncompressedSize;
      if (expandedBytes > SHAREPOINT_PACKAGE_LIMITS.expandedBytes) {
        return err(
          packageError("ArchiveExpansionTooLarge", "The expanded package exceeds safe limits.")
        );
      }
      const data = entry.getData();
      if (data.length !== uncompressedSize) {
        return err(packageError("InvalidArchive", "A package entry has an invalid size."));
      }
      entries.push({ name, data });
    }
  } catch {
    return err(packageError("InvalidArchive", "The package could not be safely read."));
  }
  return ok(entries);
}

async function parseXml(
  bytes: Buffer,
  name: string
): Promise<Result<Record<string, unknown>, SharePointPackageInspectionError>> {
  if (bytes.length > SHAREPOINT_PACKAGE_LIMITS.metadataBytes) {
    return err(packageError("MetadataTooLarge", `Package metadata "${name}" exceeds safe limits.`));
  }
  try {
    const parsed: unknown = await xml2js.parseStringPromise(bytes.toString("utf8"), {
      explicitArray: true,
      strict: true,
      trim: true,
      explicitRoot: true,
    });
    if (!isRecord(parsed)) {
      return err(packageError("InvalidPackageMetadata", `Package metadata "${name}" is invalid.`));
    }
    return ok(parsed);
  } catch {
    return err(packageError("InvalidPackageMetadata", `Package metadata "${name}" is invalid.`));
  }
}

async function validateCopilotComponents(
  entries: ArchiveEntry[]
): Promise<Result<void, SharePointPackageInspectionError>> {
  const componentManifestPattern =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}\/CopilotComponent_[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}\.xml$/i;
  const manifests = entries.filter((entry) => componentManifestPattern.test(entry.name));
  if (manifests.length === 0) {
    return err(
      packageError(
        "MissingSharePointComponent",
        "The SharePoint package must contain at least one Copilot component manifest."
      )
    );
  }
  for (const manifest of manifests) {
    const parsed = await parseXml(manifest.data, manifest.name);
    if (parsed.isErr()) {
      return err(parsed.error);
    }
    const elements = parsed.value["Elements"];
    if (
      !isRecord(elements) ||
      !Array.isArray(elements["ClientSideComponent"]) ||
      elements["ClientSideComponent"].length === 0
    ) {
      return err(
        packageError(
          "InvalidCopilotComponent",
          `Package metadata "${manifest.name}" is not a valid Copilot component manifest.`
        )
      );
    }
  }
  return ok(undefined);
}

function relationshipSourcePart(relsName: string): string | undefined {
  if (relsName === "_rels/.rels") {
    return "";
  }
  if (relsName.startsWith("_rels/") && relsName.endsWith(".rels")) {
    return relsName.slice("_rels/".length, -".rels".length);
  }
  const marker = "/_rels/";
  const markerIndex = relsName.lastIndexOf(marker);
  if (markerIndex < 0 || !relsName.endsWith(".rels")) {
    return undefined;
  }
  return `${relsName.slice(0, markerIndex)}/${relsName.slice(
    markerIndex + marker.length,
    -".rels".length
  )}`;
}

function resolveRelationshipTarget(sourcePart: string, target: string): string | undefined {
  if (target.length === 0 || target.includes("\\") || target.includes("\0")) {
    return undefined;
  }
  let decodedTarget: string;
  try {
    decodedTarget = decodeURIComponent(target.split(/[?#]/, 1)[0]);
  } catch {
    return undefined;
  }
  if (decodedTarget.startsWith("//")) {
    return undefined;
  }
  const resolved = path.posix.normalize(
    decodedTarget.startsWith("/")
      ? decodedTarget.slice(1)
      : path.posix.join(path.posix.dirname(sourcePart), decodedTarget)
  );
  return resolved === ".." || resolved.startsWith("../") ? undefined : resolved;
}

async function findEmbeddedPlugin(
  entries: ArchiveEntry[]
): Promise<
  Result<{ packagePath: string; pluginManifest: Buffer }, SharePointPackageInspectionError>
> {
  const byName = new Map(entries.map((entry) => [entry.name, entry.data]));
  const candidates: Array<{ packagePath: string; pluginManifest: Buffer }> = [];
  for (const entry of entries) {
    if (!entry.name.toLowerCase().endsWith(".zip")) {
      continue;
    }
    const nestedEntries = readBoundedArchive(entry.data);
    if (nestedEntries.isErr()) {
      return err(nestedEntries.error);
    }
    const manifests = nestedEntries.value.filter((nested) => nested.name === "ai-plugin.json");
    if (manifests.length > 1) {
      return err(
        packageError("AmbiguousPluginManifest", "The embedded agent has duplicate API plugins.")
      );
    }
    if (manifests.length === 1) {
      if (manifests[0].data.length > SHAREPOINT_PACKAGE_LIMITS.metadataBytes) {
        return err(
          packageError("MetadataTooLarge", "The embedded API plugin exceeds safe metadata limits.")
        );
      }
      candidates.push({ packagePath: entry.name, pluginManifest: manifests[0].data });
    }
  }
  if (candidates.length !== 1) {
    return err(
      packageError(
        "EmbeddedPluginNotUnique",
        "The package must contain exactly one embedded API plugin."
      )
    );
  }

  let plugin: unknown;
  try {
    plugin = JSON.parse(candidates[0].pluginManifest.toString("utf8"));
  } catch {
    return err(packageError("InvalidPluginManifest", "The embedded API plugin is invalid JSON."));
  }
  if (
    typeof plugin !== "object" ||
    plugin === null ||
    Array.isArray(plugin) ||
    !("runtimes" in plugin) ||
    !Array.isArray(plugin.runtimes) ||
    plugin.runtimes.length !== 1 ||
    typeof plugin.runtimes[0] !== "object" ||
    plugin.runtimes[0] === null ||
    !("type" in plugin.runtimes[0]) ||
    plugin.runtimes[0].type !== "RemoteMCPServer"
  ) {
    return err(
      packageError(
        "UnsupportedPluginRuntime",
        "The embedded API plugin must use the SharePoint Remote MCP runtime."
      )
    );
  }

  const relationships = entries.filter((entry) => entry.name.endsWith(".rels"));
  let linked = false;
  for (const relsEntry of relationships) {
    const sourcePart = relationshipSourcePart(relsEntry.name);
    if (sourcePart === undefined) {
      continue;
    }
    const parsed = await parseXml(relsEntry.data, relsEntry.name);
    if (parsed.isErr()) {
      return err(parsed.error);
    }
    const relationshipsRoot = parsed.value["Relationships"];
    if (!isRecord(relationshipsRoot)) {
      return err(
        packageError("InvalidPackageMetadata", `Package metadata "${relsEntry.name}" is invalid.`)
      );
    }
    const relsValue = relationshipsRoot["Relationship"];
    const rels = isRecordArray(relsValue) ? relsValue : [];
    if (
      rels.some(
        (relationship) =>
          getAttribute(relationship, "TargetMode")?.toLowerCase() !== "external" &&
          getAttribute(relationship, "Target") !== undefined &&
          resolveRelationshipTarget(sourcePart, getAttribute(relationship, "Target") ?? "") ===
            candidates[0].packagePath
      )
    ) {
      linked = true;
      break;
    }
  }
  if (!linked || !byName.has(candidates[0].packagePath)) {
    return err(
      packageError(
        "PluginPackageNotLinked",
        "The embedded API plugin is not linked by OPC metadata."
      )
    );
  }
  return ok(candidates[0]);
}

async function removeEmbeddedPackage(
  bytes: Buffer,
  entries: ArchiveEntry[],
  packagePath: string
): Promise<Result<Buffer, SharePointPackageInspectionError>> {
  const zip = new AdmZip(bytes);
  const byName = new Map(entries.map((entry) => [entry.name, entry.data]));
  const builder = new xml2js.Builder({ renderOpts: { pretty: false } });
  for (const entry of entries) {
    if (!entry.name.endsWith(".rels")) {
      continue;
    }
    const sourcePart = relationshipSourcePart(entry.name);
    if (sourcePart === undefined) {
      continue;
    }
    const parsed = await parseXml(entry.data, entry.name);
    if (parsed.isErr()) {
      return err(parsed.error);
    }
    const relationships = parsed.value["Relationships"];
    if (relationships === undefined) {
      return err(
        packageError("InvalidPackageMetadata", `Package metadata "${entry.name}" is invalid.`)
      );
    }
    if (!isRecord(relationships)) {
      return err(
        packageError("InvalidPackageMetadata", `Package metadata "${entry.name}" is invalid.`)
      );
    }
    const relationshipValue = relationships["Relationship"];
    const existing = isRecordArray(relationshipValue) ? relationshipValue : [];
    relationships["Relationship"] = existing.filter(
      (relationship) =>
        getAttribute(relationship, "TargetMode")?.toLowerCase() === "external" ||
        getAttribute(relationship, "Target") === undefined ||
        resolveRelationshipTarget(sourcePart, getAttribute(relationship, "Target") ?? "") !==
          packagePath
    );
    const filtered = relationships["Relationship"];
    if (Array.isArray(filtered) && filtered.length !== existing.length) {
      zip.updateFile(entry.name, Buffer.from(builder.buildObject(parsed.value), "utf8"));
    }
  }

  const contentTypes = byName.get("[Content_Types].xml");
  if (contentTypes === undefined) {
    return err(
      packageError("MissingContentTypes", "The SharePoint package is missing OPC content types.")
    );
  }
  const parsedContentTypes = await parseXml(contentTypes, "[Content_Types].xml");
  if (parsedContentTypes.isErr()) {
    return err(parsedContentTypes.error);
  }
  const types = parsedContentTypes.value["Types"];
  if (!isRecord(types)) {
    return err(
      packageError("InvalidContentTypes", "The SharePoint package content types are invalid.")
    );
  }
  const overrides = types["Override"];
  types["Override"] = (isRecordArray(overrides) ? overrides : []).filter((override) => {
    const partName = getAttribute(override, "PartName");
    if (partName === undefined) {
      return true;
    }
    return resolveRelationshipTarget("", partName) !== packagePath;
  });
  const defaults = types["Default"];
  const remainingZip = entries.some(
    (entry) => entry.name !== packagePath && entry.name.toLowerCase().endsWith(".zip")
  );
  if (!remainingZip) {
    types["Default"] = (isRecordArray(defaults) ? defaults : []).filter(
      (value) => getAttribute(value, "Extension")?.toLowerCase() !== "zip"
    );
  }
  zip.updateFile(
    "[Content_Types].xml",
    Buffer.from(builder.buildObject(parsedContentTypes.value), "utf8")
  );
  zip.deleteFile(packagePath);
  try {
    return ok(zip.toBuffer());
  } catch {
    return err(
      packageError(
        "PackageTransformFailed",
        "The SharePoint deployment package could not be prepared."
      )
    );
  }
}

export async function inspectSharePointPackage(
  bytes: Buffer
): Promise<Result<SharePointPackageInspection, SharePointPackageInspectionError>> {
  const readResult = readBoundedArchive(bytes);
  if (readResult.isErr()) {
    return err(readResult.error);
  }
  const entries = readResult.value;
  const componentResult = await validateCopilotComponents(entries);
  if (componentResult.isErr()) {
    return err(componentResult.error);
  }
  const entryMap = new Map(entries.map((entry) => [entry.name, entry.data]));
  const identity = await readSharePointPackageIdentity(entryMap);
  if (identity.isErr()) {
    return err(identity.error);
  }
  const pluginResult = await findEmbeddedPlugin(entries);
  if (pluginResult.isErr()) {
    return err(pluginResult.error);
  }
  const deploymentResult = await removeEmbeddedPackage(
    bytes,
    entries,
    pluginResult.value.packagePath
  );
  if (deploymentResult.isErr()) {
    return err(deploymentResult.error);
  }
  return ok({
    sourceSha256: sha256(bytes),
    deploymentBytes: deploymentResult.value,
    deploymentSha256: sha256(deploymentResult.value),
    productId: identity.value.productId,
    version: identity.value.version,
    skipFeatureDeployment: identity.value.skipFeatureDeployment,
    embeddedPackagePath: pluginResult.value.packagePath,
    pluginManifest: pluginResult.value.pluginManifest,
  });
}

export async function inspectSharePointPackageIdentity(
  bytes: Buffer
): Promise<Result<SharePointPackageIdentity, SharePointPackageInspectionError>> {
  const readResult = readBoundedArchive(bytes);
  if (readResult.isErr()) {
    return err(readResult.error);
  }
  return readSharePointPackageIdentity(
    new Map(readResult.value.map((entry) => [entry.name, entry.data]))
  );
}

async function readSharePointPackageIdentity(
  entries: ReadonlyMap<string, Buffer>
): Promise<Result<SharePointPackageIdentity, SharePointPackageInspectionError>> {
  const manifestBytes = entries.get("AppManifest.xml");
  if (manifestBytes === undefined) {
    return err(
      packageError("MissingAppManifest", "The SharePoint package is missing AppManifest.xml.")
    );
  }
  const appManifest = await parseXml(manifestBytes, "AppManifest.xml");
  if (appManifest.isErr()) {
    return err(appManifest.error);
  }
  const app = appManifest.value["App"];
  const productId = isRecord(app)
    ? getAttribute(app, "ProductID")?.replace(/[{}]/g, "")
    : undefined;
  const version = isRecord(app) ? getAttribute(app, "Version") : undefined;
  const skipFeatureDeployment = isRecord(app)
    ? getAttribute(app, "SkipFeatureDeployment")
    : undefined;
  if (
    productId === undefined ||
    !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(productId) ||
    version === undefined ||
    !/^\d+\.\d+\.\d+\.\d+$/.test(version) ||
    (skipFeatureDeployment !== "true" && skipFeatureDeployment !== "false")
  ) {
    return err(
      packageError("InvalidAppManifest", "The SharePoint package identity or version is invalid.")
    );
  }
  return ok({
    productId,
    version,
    skipFeatureDeployment: skipFeatureDeployment === "true",
  });
}

export async function readSharePointPackage(
  packagePath: string
): Promise<Result<SharePointPackageInspection, SharePointPackageInspectionError>> {
  let fileDescriptor: number | undefined;
  try {
    fileDescriptor = await fs.open(packagePath, "r");
    const stats = await fs.fstat(fileDescriptor);
    if (stats.size > SHAREPOINT_PACKAGE_LIMITS.inputBytes) {
      return err(packageError("ArchiveTooLarge", "The package exceeds the supported input size."));
    }
    const bytes = await fs.readFile(fileDescriptor);
    if (bytes.length > SHAREPOINT_PACKAGE_LIMITS.inputBytes) {
      return err(packageError("ArchiveTooLarge", "The package exceeds the supported input size."));
    }
    return await inspectSharePointPackage(bytes);
  } catch {
    return err(
      packageError(
        "PackageReadFailed",
        `The SharePoint package "${path.basename(packagePath)}" could not be read.`
      )
    );
  } finally {
    if (fileDescriptor !== undefined) {
      await fs.close(fileDescriptor);
    }
  }
}
