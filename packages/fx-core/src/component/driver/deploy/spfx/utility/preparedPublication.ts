// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { createHash, randomUUID } from "crypto";
import fs from "fs-extra";
import path from "path";
import properLock from "proper-lockfile";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";
import {
  inspectSharePointPackage,
  inspectSharePointPackageIdentity,
  readBoundedArchive,
  SHAREPOINT_PACKAGE_LIMITS,
  SharePointPackageInspectionError,
} from "./packageFile";

const journalVersion = 1;
const publicationDirectory = path.join(".fx", "sharepoint-publications");
const preparedPluginPlaceholder = /\{\{\s*([A-Z][A-Z0-9_]*)\s*\}\}/g;
const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export type PreparedPublicationState =
  | "prepared"
  | "sharePointUploadInFlight"
  | "sharePointUploaded"
  | "sharePointDeployInFlight"
  | "sharePointVerified"
  | "appSubmissionInFlight"
  | "submitted"
  | "rejected"
  | "unknown";

export type PreparedWriteState = "notStarted" | "inFlight" | "confirmed" | "rejected" | "unknown";

export interface PreparedPublicationJournal {
  journalVersion: 1;
  publicationId: string;
  state: PreparedPublicationState;
  sourceSharePoint: {
    path: string;
    sha256: string;
    productId: string;
    version: string;
    skipFeatureDeployment: boolean;
  };
  sourceApplication: {
    path: string;
    sha256: string;
    appId: string;
    version: string;
  };
  plugin: {
    sha256: string;
  };
  sharePointArtifact: {
    path: string;
    sha256: string;
    uploadState: PreparedWriteState;
    deployState: PreparedWriteState;
    catalogUrl?: string;
    fileName?: string;
  };
  applicationArtifact: {
    path: string;
    sha256: string;
    appId: string;
    version: string;
    submissionState: PreparedWriteState;
    submittedAppId?: string;
  };
}

export class PreparedPublicationError extends Error {
  public constructor(
    public readonly code: string,
    message: string,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = "PreparedPublicationError";
  }
}

export function isPreparedCatalogUrl(value: string, tenantUrl: string): boolean {
  try {
    const catalog = new URL(value);
    const tenant = new URL(tenantUrl);
    return (
      catalog.protocol === "https:" &&
      catalog.origin === tenant.origin &&
      !catalog.username &&
      !catalog.password &&
      !catalog.search &&
      !catalog.hash
    );
  } catch {
    return false;
  }
}

export function compareSharePointPackageVersions(
  left: string,
  right: string
): -1 | 0 | 1 | undefined {
  const versionPattern = /^\d+\.\d+\.\d+\.\d+$/;
  if (!versionPattern.test(left) || !versionPattern.test(right)) {
    return undefined;
  }
  const leftParts = left.split(".").map((part) => BigInt(part));
  const rightParts = right.split(".").map((part) => BigInt(part));
  for (let index = 0; index < leftParts.length; index += 1) {
    if (leftParts[index] > rightParts[index]) {
      return 1;
    }
    if (leftParts[index] < rightParts[index]) {
      return -1;
    }
  }
  return 0;
}

export interface PreparePublicationInput {
  projectPath: string;
  packagePath: string;
  appPackagePath: string;
}

export interface PreparedPublicationOutput {
  publicationId: string;
  preparedPublicationPath: string;
  sourcePackageSha256: string;
  sourceAppPackageSha256: string;
  pluginSha256: string;
  sharePointPackagePath: string;
  sharePointPackageSha256: string;
  appPackagePath: string;
  appPackageSha256: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function sha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function projectRelative(root: string, filePath: string): string {
  return path.relative(root, filePath).split(path.sep).join("/");
}

function isWithinProject(root: string, filePath: string): boolean {
  const relative = path.relative(root, filePath);
  return relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function resolveJournalPath(
  root: string,
  publicationId: string,
  journalPath: string
): string | undefined {
  if (!guidPattern.test(publicationId) || path.isAbsolute(journalPath)) {
    return undefined;
  }
  const expected = path.resolve(root, publicationDirectory, publicationId, "journal.json");
  const provided = path.resolve(root, journalPath);
  return provided === expected && isWithinProject(root, provided) ? expected : undefined;
}

async function readFileBounded(
  filePath: string,
  maxBytes: number
): Promise<Result<Buffer, PreparedPublicationError>> {
  let descriptor: number | undefined;
  try {
    descriptor = await fs.open(filePath, "r");
    const stats = await fs.fstat(descriptor);
    if (stats.size > maxBytes) {
      await fs.close(descriptor);
      descriptor = undefined;
      return err(
        new PreparedPublicationError(
          "InputTooLarge",
          "A publication input exceeds safe size limits."
        )
      );
    }
    const bytes = await fs.readFile(descriptor);
    await fs.close(descriptor);
    descriptor = undefined;
    if (bytes.length > maxBytes) {
      return err(
        new PreparedPublicationError(
          "InputTooLarge",
          "A publication input exceeds safe size limits."
        )
      );
    }
    return ok(bytes);
  } catch (error) {
    return err(
      new PreparedPublicationError("InputReadFailed", "A publication input could not be read.", {
        cause: error,
      })
    );
  } finally {
    if (descriptor !== undefined) {
      await fs.close(descriptor);
    }
  }
}

async function readProjectInput(
  root: string,
  inputPath: string
): Promise<
  Result<{ path: string; relativePath: string; bytes: Buffer }, PreparedPublicationError>
> {
  if (path.isAbsolute(inputPath)) {
    return err(
      new PreparedPublicationError(
        "InvalidInputPath",
        "Publication inputs must be project-relative."
      )
    );
  }
  const candidate = path.resolve(root, inputPath);
  if (!isWithinProject(root, candidate)) {
    return err(
      new PreparedPublicationError(
        "InvalidInputPath",
        "Publication inputs must remain inside the project."
      )
    );
  }
  let resolved: string;
  try {
    resolved = await fs.realpath(candidate);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InputReadFailed",
        "A publication input could not be resolved.",
        {
          cause: error,
        }
      )
    );
  }
  if (!isWithinProject(root, resolved)) {
    return err(
      new PreparedPublicationError(
        "InvalidInputPath",
        "Publication inputs must remain inside the project."
      )
    );
  }
  const read = await readFileBounded(resolved, SHAREPOINT_PACKAGE_LIMITS.inputBytes);
  if (read.isErr()) {
    return err(read.error);
  }
  return ok({ path: resolved, relativePath: projectRelative(root, resolved), bytes: read.value });
}

async function findUnresolvedPublication(
  root: string,
  sourceSharePointSha256: string,
  sourceApplicationSha256: string
): Promise<Result<void, PreparedPublicationError>> {
  const directory = path.resolve(root, publicationDirectory);
  let entries: fs.Dirent[];
  try {
    entries = await fs.readdir(directory, { withFileTypes: true });
  } catch (error) {
    if (isRecord(error) && error["code"] === "ENOENT") {
      return ok(undefined);
    }
    return err(
      new PreparedPublicationError(
        "PublicationRecoveryCheckFailed",
        "Existing publication state could not be checked safely.",
        { cause: error }
      )
    );
  }
  const publicationDirectories = entries.filter(
    (entry) => entry.isDirectory() && guidPattern.test(entry.name)
  );
  if (publicationDirectories.length > SHAREPOINT_PACKAGE_LIMITS.entries) {
    return err(
      new PreparedPublicationError(
        "TooManyPublicationRecords",
        "The project contains too many prepared publication records to check safely."
      )
    );
  }
  for (const entry of publicationDirectories) {
    const journalPath = projectRelative(root, path.join(directory, entry.name, "journal.json"));
    const loaded = await loadPreparedPublication(root, entry.name, journalPath);
    if (loaded.isErr()) {
      return err(
        new PreparedPublicationError(
          "PublicationRecoveryCheckFailed",
          "An existing publication record could not be checked safely.",
          { cause: loaded.error }
        )
      );
    }
    const journal = loaded.value.journal;
    if (
      journal.sourceSharePoint.sha256 === sourceSharePointSha256 &&
      journal.sourceApplication.sha256 === sourceApplicationSha256 &&
      (isAmbiguousPublicationState(journal) || journal.state === "rejected")
    ) {
      return err(
        new PreparedPublicationError(
          "UnresolvedPreparedPublication",
          "A previous publication of these exact source packages has a rejected or uncertain remote outcome and requires operator reconciliation before retry."
        )
      );
    }
  }
  return ok(undefined);
}

function replaceTenantValues(pluginBytes: Buffer): Result<Buffer, PreparedPublicationError> {
  const source = pluginBytes.toString("utf8");
  const placeholders = [...source.matchAll(preparedPluginPlaceholder)].map((match) => match[1]);
  const distinct = [...new Set(placeholders)];
  let output = source;
  for (const placeholder of distinct) {
    if (placeholder !== "TENANT_MCP_URL") {
      return err(
        new PreparedPublicationError(
          "UnsupportedPluginPlaceholder",
          `The embedded API plugin contains unsupported placeholder "${placeholder}".`
        )
      );
    }
    const value = process.env[placeholder];
    if (!value) {
      return err(
        new PreparedPublicationError(
          "MissingTenantPluginValue",
          `The required environment variable "${placeholder}" is not set.`
        )
      );
    }
    let url: URL;
    try {
      url = new URL(value);
    } catch (error) {
      return err(
        new PreparedPublicationError("InvalidTenantPluginValue", "The tenant API URL is invalid.", {
          cause: error,
        })
      );
    }
    if (url.protocol !== "https:" || url.username || url.password || url.hash) {
      return err(
        new PreparedPublicationError(
          "InvalidTenantPluginValue",
          "The tenant API URL must be an HTTPS URL without embedded credentials or a fragment."
        )
      );
    }
    output = output.replace(
      new RegExp(`\\{\\{\\s*${placeholder}\\s*\\}\\}`, "g"),
      url.toString().replace(/\/$/, "")
    );
  }
  if (/\{\{[^{}]+\}\}/.test(output)) {
    return err(
      new PreparedPublicationError(
        "UnresolvedPluginPlaceholder",
        "The embedded API plugin contains an unresolved placeholder."
      )
    );
  }
  try {
    JSON.parse(output);
    return ok(Buffer.from(output, "utf8"));
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InvalidPluginManifest",
        "The resolved API plugin is invalid JSON.",
        {
          cause: error,
        }
      )
    );
  }
}

function jsonObject(
  bytes: Buffer,
  name: string
): Result<Record<string, unknown>, PreparedPublicationError> {
  if (bytes.length > SHAREPOINT_PACKAGE_LIMITS.metadataBytes) {
    return err(
      new PreparedPublicationError(
        "MetadataTooLarge",
        `Package metadata "${name}" exceeds safe limits.`
      )
    );
  }
  try {
    const parsed: unknown = JSON.parse(bytes.toString("utf8").replace(/^\uFEFF/, ""));
    if (!isRecord(parsed)) {
      return err(
        new PreparedPublicationError(
          "InvalidPackageMetadata",
          `Package metadata "${name}" is invalid.`
        )
      );
    }
    return ok(parsed);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InvalidPackageMetadata",
        `Package metadata "${name}" is invalid.`,
        {
          cause: error,
        }
      )
    );
  }
}

function getAgentReference(manifest: Record<string, unknown>): Result<
  {
    container: Record<string, unknown>;
    property: "declarativeAgents" | "declarativeCopilots";
    reference: Record<string, unknown>;
  },
  PreparedPublicationError
> {
  const references: Array<{
    container: Record<string, unknown>;
    property: "declarativeAgents" | "declarativeCopilots";
    reference: Record<string, unknown>;
  }> = [];
  const referenceTypes: {
    containerName: string;
    property: "declarativeAgents" | "declarativeCopilots";
  }[] = [
    { containerName: "copilotAgents", property: "declarativeAgents" },
    { containerName: "copilotExtensions", property: "declarativeCopilots" },
  ];
  for (const referenceType of referenceTypes) {
    const container = manifest[referenceType.containerName];
    if (!isRecord(container)) {
      continue;
    }
    const values = container[referenceType.property];
    if (!Array.isArray(values)) {
      continue;
    }
    for (const value of values) {
      if (isRecord(value)) {
        references.push({
          container,
          property: referenceType.property,
          reference: value,
        });
      }
    }
  }
  if (references.length !== 1) {
    return err(
      new PreparedPublicationError(
        "UnsupportedApplicationPackage",
        "The application package must reference exactly one declarative agent."
      )
    );
  }
  return ok(references[0]);
}

function prepareApplicationPackage(
  bytes: Buffer,
  pluginBytes: Buffer,
  pluginHash: string
): Result<{ bytes: Buffer; appId: string; version: string }, PreparedPublicationError> {
  const entries = readBoundedArchive(bytes);
  if (entries.isErr()) {
    return err(
      new PreparedPublicationError(
        "InvalidApplicationPackage",
        "The application package is invalid.",
        {
          cause: entries.error,
        }
      )
    );
  }
  const byName = new Map(entries.value.map((entry) => [entry.name, entry.data]));
  const manifestBytes = byName.get("manifest.json");
  if (manifestBytes === undefined) {
    return err(
      new PreparedPublicationError(
        "InvalidApplicationPackage",
        "The application package is missing manifest.json."
      )
    );
  }
  const manifestResult = jsonObject(manifestBytes, "manifest.json");
  if (manifestResult.isErr()) {
    return err(manifestResult.error);
  }
  const manifest = manifestResult.value;
  const appId = manifest["id"];
  const version = manifest["version"];
  if (
    typeof appId !== "string" ||
    !guidPattern.test(appId) ||
    typeof version !== "string" ||
    !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(version)
  ) {
    return err(
      new PreparedPublicationError(
        "InvalidApplicationIdentity",
        "The application identity or version is invalid."
      )
    );
  }
  const agentReference = getAgentReference(manifest);
  if (agentReference.isErr()) {
    return err(agentReference.error);
  }
  const agentFile = agentReference.value.reference["file"];
  if (typeof agentFile !== "string" || path.posix.normalize(agentFile) !== agentFile) {
    return err(
      new PreparedPublicationError(
        "InvalidAgentReference",
        "The application agent reference is invalid."
      )
    );
  }
  const agentBytes = byName.get(agentFile);
  if (agentBytes === undefined) {
    return err(
      new PreparedPublicationError(
        "MissingAgentManifest",
        "The application package is missing its declarative agent manifest."
      )
    );
  }
  const agentResult = jsonObject(agentBytes, agentFile);
  if (agentResult.isErr()) {
    return err(agentResult.error);
  }
  const agent = agentResult.value;
  const actionsValue = agent["actions"];
  if (actionsValue !== undefined && !Array.isArray(actionsValue)) {
    return err(
      new PreparedPublicationError(
        "InvalidAgentActions",
        "The declarative agent actions must be an array."
      )
    );
  }
  const actions = Array.isArray(actionsValue) ? actionsValue.filter(isRecord) : [];
  if (actionsValue !== undefined && actions.length !== actionsValue.length) {
    return err(
      new PreparedPublicationError(
        "InvalidAgentActions",
        "The declarative agent actions are invalid."
      )
    );
  }
  const pluginFile = `plugins/spfx-api-${pluginHash.slice(0, 16)}.json`;
  const pluginActionId = `spfx-api-${pluginHash.slice(0, 16)}`;
  const existingPlugin = byName.get(pluginFile);
  if (existingPlugin !== undefined && !existingPlugin.equals(pluginBytes)) {
    return err(
      new PreparedPublicationError(
        "ApplicationPackageCollision",
        "The application package contains conflicting plugin content."
      )
    );
  }
  if (
    actions.some(
      (action) =>
        action["file"] === pluginFile ||
        (action["id"] === pluginActionId && action["file"] !== pluginFile)
    )
  ) {
    return err(
      new PreparedPublicationError(
        "ApplicationPackageCollision",
        "The application package already references this prepared plugin."
      )
    );
  }
  agent["actions"] = [
    ...actions,
    {
      id: pluginActionId,
      file: pluginFile,
    },
  ];

  const zip = new AdmZip(bytes);
  zip.updateFile(agentFile, Buffer.from(JSON.stringify(agent, null, 2), "utf8"));
  if (existingPlugin === undefined) {
    zip.addFile(pluginFile, pluginBytes);
  }
  try {
    const output = zip.toBuffer();
    const outputEntries = readBoundedArchive(output);
    if (outputEntries.isErr()) {
      return err(outputEntries.error);
    }
    return ok({ bytes: output, appId, version });
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "ApplicationPackagePrepareFailed",
        "The application package could not be prepared.",
        {
          cause: error,
        }
      )
    );
  }
}

async function atomicWrite(filePath: string, bytes: Buffer): Promise<void> {
  const tempPath = `${filePath}.${randomUUID()}.tmp`;
  try {
    await fs.writeFile(tempPath, bytes, { flag: "wx" });
    await fs.rename(tempPath, filePath);
  } catch (error) {
    await fs.remove(tempPath);
    throw new PreparedPublicationError(
      "PublicationWriteFailed",
      "Prepared publication data could not be persisted.",
      {
        cause: error,
      }
    );
  }
}

export async function preparePublication(
  input: PreparePublicationInput
): Promise<
  Result<PreparedPublicationOutput, PreparedPublicationError | SharePointPackageInspectionError>
> {
  let root: string;
  try {
    root = await fs.realpath(input.projectPath);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "ProjectPathUnavailable",
        "The project directory could not be resolved.",
        {
          cause: error,
        }
      )
    );
  }
  const sourcePackage = await readProjectInput(root, input.packagePath);
  if (sourcePackage.isErr()) {
    return err(sourcePackage.error);
  }
  if (!/^[A-Za-z0-9._-]+\.sppkg$/i.test(path.posix.basename(sourcePackage.value.relativePath))) {
    return err(
      new PreparedPublicationError(
        "UnsupportedSharePointPackageFileName",
        "The SharePoint package file name must use only letters, numbers, dots, underscores, or hyphens."
      )
    );
  }
  const sourceAppPackage = await readProjectInput(root, input.appPackagePath);
  if (sourceAppPackage.isErr()) {
    return err(sourceAppPackage.error);
  }
  const recoveryCheck = await findUnresolvedPublication(
    root,
    sha256(sourcePackage.value.bytes),
    sha256(sourceAppPackage.value.bytes)
  );
  if (recoveryCheck.isErr()) {
    return err(recoveryCheck.error);
  }
  const sharePointPackage = await inspectSharePointPackage(sourcePackage.value.bytes);
  if (sharePointPackage.isErr()) {
    return err(sharePointPackage.error);
  }
  const resolvedPlugin = replaceTenantValues(sharePointPackage.value.pluginManifest);
  if (resolvedPlugin.isErr()) {
    return err(resolvedPlugin.error);
  }
  const pluginSha256 = sha256(resolvedPlugin.value);
  const appPackage = prepareApplicationPackage(
    sourceAppPackage.value.bytes,
    resolvedPlugin.value,
    pluginSha256
  );
  if (appPackage.isErr()) {
    return err(appPackage.error);
  }

  const publicationId = randomUUID();
  const stagingDirectory = path.resolve(root, publicationDirectory, publicationId);
  if (!isWithinProject(root, stagingDirectory)) {
    return err(
      new PreparedPublicationError("InvalidStagingPath", "The publication staging path is invalid.")
    );
  }
  try {
    await fs.ensureDir(stagingDirectory);
    const resolvedStagingDirectory = await fs.realpath(stagingDirectory);
    if (!isWithinProject(root, resolvedStagingDirectory)) {
      return err(
        new PreparedPublicationError(
          "InvalidStagingPath",
          "The publication staging path is invalid."
        )
      );
    }
    const sharePointArtifactPath = path.join(resolvedStagingDirectory, "sharepoint.sppkg");
    const applicationArtifactPath = path.join(resolvedStagingDirectory, "application.zip");
    const journalPath = path.join(resolvedStagingDirectory, "journal.json");
    await atomicWrite(sharePointArtifactPath, sharePointPackage.value.deploymentBytes);
    await atomicWrite(applicationArtifactPath, appPackage.value.bytes);
    const journal: PreparedPublicationJournal = {
      journalVersion,
      publicationId,
      state: "prepared",
      sourceSharePoint: {
        path: sourcePackage.value.relativePath,
        sha256: sha256(sourcePackage.value.bytes),
        productId: sharePointPackage.value.productId,
        version: sharePointPackage.value.version,
        skipFeatureDeployment: sharePointPackage.value.skipFeatureDeployment,
      },
      sourceApplication: {
        path: sourceAppPackage.value.relativePath,
        sha256: sha256(sourceAppPackage.value.bytes),
        appId: appPackage.value.appId,
        version: appPackage.value.version,
      },
      plugin: { sha256: pluginSha256 },
      sharePointArtifact: {
        path: projectRelative(root, sharePointArtifactPath),
        sha256: sharePointPackage.value.deploymentSha256,
        uploadState: "notStarted",
        deployState: "notStarted",
      },
      applicationArtifact: {
        path: projectRelative(root, applicationArtifactPath),
        sha256: sha256(appPackage.value.bytes),
        appId: appPackage.value.appId,
        version: appPackage.value.version,
        submissionState: "notStarted",
      },
    };
    await atomicWrite(journalPath, Buffer.from(JSON.stringify(journal, null, 2), "utf8"));
    return ok({
      publicationId,
      preparedPublicationPath: projectRelative(root, journalPath),
      sourcePackageSha256: journal.sourceSharePoint.sha256,
      sourceAppPackageSha256: journal.sourceApplication.sha256,
      pluginSha256,
      sharePointPackagePath: journal.sharePointArtifact.path,
      sharePointPackageSha256: journal.sharePointArtifact.sha256,
      appPackagePath: journal.applicationArtifact.path,
      appPackageSha256: journal.applicationArtifact.sha256,
    });
  } catch (error) {
    await fs.remove(stagingDirectory);
    return err(
      error instanceof PreparedPublicationError
        ? error
        : new PreparedPublicationError(
            "PublicationPrepareFailed",
            "The publication artifacts could not be prepared.",
            {
              cause: error,
            }
          )
    );
  }
}

function isPublicationState(value: unknown): value is PreparedPublicationState {
  return (
    typeof value === "string" &&
    [
      "prepared",
      "sharePointUploadInFlight",
      "sharePointUploaded",
      "sharePointDeployInFlight",
      "sharePointVerified",
      "appSubmissionInFlight",
      "submitted",
      "rejected",
      "unknown",
    ].includes(value)
  );
}

function isWriteState(value: unknown): value is PreparedWriteState {
  return (
    typeof value === "string" &&
    ["notStarted", "inFlight", "confirmed", "rejected", "unknown"].includes(value)
  );
}

function isArtifactRecord(value: unknown): value is Record<string, unknown> {
  return (
    isRecord(value) &&
    isSafeRelativePath(value["path"]) &&
    typeof value["sha256"] === "string" &&
    /^[0-9a-f]{64}$/i.test(value["sha256"])
  );
}

function isSafeRelativePath(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length > 0 &&
    !value.startsWith("/") &&
    !value.includes("\\") &&
    path.posix.normalize(value) === value &&
    value.split("/").every((segment) => segment !== ".." && segment !== "." && segment.length > 0)
  );
}

function parseJournal(value: unknown): PreparedPublicationJournal | undefined {
  if (
    !isRecord(value) ||
    value["journalVersion"] !== journalVersion ||
    typeof value["publicationId"] !== "string" ||
    !guidPattern.test(value["publicationId"]) ||
    !isPublicationState(value["state"]) ||
    !isArtifactRecord(value["sourceSharePoint"]) ||
    typeof value["sourceSharePoint"]["productId"] !== "string" ||
    typeof value["sourceSharePoint"]["version"] !== "string" ||
    typeof value["sourceSharePoint"]["skipFeatureDeployment"] !== "boolean" ||
    !isArtifactRecord(value["sourceApplication"]) ||
    typeof value["sourceApplication"]["appId"] !== "string" ||
    typeof value["sourceApplication"]["version"] !== "string" ||
    !isRecord(value["plugin"]) ||
    typeof value["plugin"]["sha256"] !== "string" ||
    !/^[0-9a-f]{64}$/i.test(value["plugin"]["sha256"]) ||
    !isArtifactRecord(value["sharePointArtifact"]) ||
    !isWriteState(value["sharePointArtifact"]["uploadState"]) ||
    !isWriteState(value["sharePointArtifact"]["deployState"]) ||
    !isArtifactRecord(value["applicationArtifact"]) ||
    typeof value["applicationArtifact"]["appId"] !== "string" ||
    typeof value["applicationArtifact"]["version"] !== "string" ||
    !isWriteState(value["applicationArtifact"]["submissionState"])
  ) {
    return undefined;
  }
  const sourceSharePoint = value["sourceSharePoint"];
  const sourceApplication = value["sourceApplication"];
  const plugin = value["plugin"];
  const sharePointArtifact = value["sharePointArtifact"];
  const applicationArtifact = value["applicationArtifact"];
  const sourceSharePointPath = sourceSharePoint["path"];
  const sourceApplicationPath = sourceApplication["path"];
  const sharePointArtifactPath = sharePointArtifact["path"];
  const applicationArtifactPath = applicationArtifact["path"];
  const sourceSharePointSha256 = sourceSharePoint["sha256"];
  const sourceSharePointProductId = sourceSharePoint["productId"];
  const sourceSharePointVersion = sourceSharePoint["version"];
  const sourceSharePointSkipFeatureDeployment = sourceSharePoint["skipFeatureDeployment"];
  const sourceApplicationSha256 = sourceApplication["sha256"];
  const sourceApplicationAppId = sourceApplication["appId"];
  const sourceApplicationVersion = sourceApplication["version"];
  const pluginSha256 = plugin["sha256"];
  const sharePointArtifactSha256 = sharePointArtifact["sha256"];
  const sharePointUploadState = sharePointArtifact["uploadState"];
  const sharePointDeployState = sharePointArtifact["deployState"];
  const applicationArtifactSha256 = applicationArtifact["sha256"];
  const applicationArtifactAppId = applicationArtifact["appId"];
  const applicationArtifactVersion = applicationArtifact["version"];
  const applicationSubmissionState = applicationArtifact["submissionState"];
  if (
    typeof sourceSharePointPath !== "string" ||
    typeof sourceApplicationPath !== "string" ||
    typeof sharePointArtifactPath !== "string" ||
    typeof applicationArtifactPath !== "string" ||
    typeof sourceSharePointSha256 !== "string" ||
    typeof sourceSharePointProductId !== "string" ||
    typeof sourceSharePointVersion !== "string" ||
    typeof sourceSharePointSkipFeatureDeployment !== "boolean" ||
    typeof sourceApplicationSha256 !== "string" ||
    typeof sourceApplicationAppId !== "string" ||
    typeof sourceApplicationVersion !== "string" ||
    typeof pluginSha256 !== "string" ||
    typeof sharePointArtifactSha256 !== "string" ||
    !isWriteState(sharePointUploadState) ||
    !isWriteState(sharePointDeployState) ||
    typeof applicationArtifactSha256 !== "string" ||
    typeof applicationArtifactAppId !== "string" ||
    typeof applicationArtifactVersion !== "string" ||
    !isWriteState(applicationSubmissionState)
  ) {
    return undefined;
  }
  return {
    journalVersion: 1,
    publicationId: value["publicationId"],
    state: value["state"],
    sourceSharePoint: {
      path: sourceSharePointPath,
      sha256: sourceSharePointSha256,
      productId: sourceSharePointProductId,
      version: sourceSharePointVersion,
      skipFeatureDeployment: sourceSharePointSkipFeatureDeployment,
    },
    sourceApplication: {
      path: sourceApplicationPath,
      sha256: sourceApplicationSha256,
      appId: sourceApplicationAppId,
      version: sourceApplicationVersion,
    },
    plugin: { sha256: pluginSha256 },
    sharePointArtifact: {
      path: sharePointArtifactPath,
      sha256: sharePointArtifactSha256,
      uploadState: sharePointUploadState,
      deployState: sharePointDeployState,
      ...(typeof sharePointArtifact["catalogUrl"] === "string"
        ? { catalogUrl: sharePointArtifact["catalogUrl"] }
        : {}),
      ...(typeof sharePointArtifact["fileName"] === "string"
        ? { fileName: sharePointArtifact["fileName"] }
        : {}),
    },
    applicationArtifact: {
      path: applicationArtifactPath,
      sha256: applicationArtifactSha256,
      appId: applicationArtifactAppId,
      version: applicationArtifactVersion,
      submissionState: applicationSubmissionState,
      ...(typeof applicationArtifact["submittedAppId"] === "string"
        ? { submittedAppId: applicationArtifact["submittedAppId"] }
        : {}),
    },
  };
}

export async function loadPreparedPublication(
  projectPath: string,
  publicationId: string,
  journalPath: string
): Promise<
  Result<
    { root: string; path: string; journal: PreparedPublicationJournal },
    PreparedPublicationError
  >
> {
  let root: string;
  try {
    root = await fs.realpath(projectPath);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "ProjectPathUnavailable",
        "The project directory could not be resolved.",
        { cause: error }
      )
    );
  }
  const resolvedPath = resolveJournalPath(root, publicationId, journalPath);
  if (resolvedPath === undefined) {
    return err(
      new PreparedPublicationError(
        "InvalidJournalPath",
        "The prepared publication journal path is invalid."
      )
    );
  }
  let realPath: string;
  try {
    realPath = await fs.realpath(resolvedPath);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "JournalReadFailed",
        "The prepared publication journal could not be resolved.",
        { cause: error }
      )
    );
  }
  if (!isWithinProject(root, realPath) || realPath !== resolvedPath) {
    return err(
      new PreparedPublicationError(
        "InvalidJournalPath",
        "The prepared publication journal path is invalid."
      )
    );
  }
  const journalBytes = await readFileBounded(realPath, SHAREPOINT_PACKAGE_LIMITS.metadataBytes);
  if (journalBytes.isErr()) {
    return err(
      new PreparedPublicationError(
        "JournalReadFailed",
        "The prepared publication journal could not be read.",
        { cause: journalBytes.error }
      )
    );
  }
  let raw: unknown;
  try {
    raw = JSON.parse(journalBytes.value.toString("utf8"));
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "JournalReadFailed",
        "The prepared publication journal could not be read.",
        { cause: error }
      )
    );
  }
  const journal = parseJournal(raw);
  if (journal === undefined || journal.publicationId !== publicationId) {
    return err(
      new PreparedPublicationError("InvalidJournal", "The prepared publication journal is invalid.")
    );
  }
  return ok({ root, path: resolvedPath, journal });
}

export async function updatePreparedPublication(
  projectPath: string,
  publicationId: string,
  journalPath: string,
  update: (journal: PreparedPublicationJournal) => PreparedPublicationJournal
): Promise<Result<PreparedPublicationJournal, PreparedPublicationError>> {
  const loaded = await loadPreparedPublication(projectPath, publicationId, journalPath);
  if (loaded.isErr()) {
    return err(loaded.error);
  }
  let release: (() => Promise<void>) | undefined;
  try {
    release = await properLock.lock(loaded.value.path, { retries: 0 });
    const reread = await loadPreparedPublication(projectPath, publicationId, journalPath);
    if (reread.isErr()) {
      return err(reread.error);
    }
    const next = update(reread.value.journal);
    await atomicWrite(loaded.value.path, Buffer.from(JSON.stringify(next, null, 2), "utf8"));
    return ok(next);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "JournalUpdateFailed",
        "The prepared publication journal could not be updated.",
        {
          cause: error,
        }
      )
    );
  } finally {
    if (release !== undefined) {
      await release();
    }
  }
}

export async function verifyPreparedArtifact(
  root: string,
  artifact: { path: string; sha256: string }
): Promise<Result<Buffer, PreparedPublicationError>> {
  const fullPath = path.resolve(root, ...artifact.path.split("/"));
  if (!isWithinProject(root, fullPath)) {
    return err(
      new PreparedPublicationError("InvalidArtifactPath", "A prepared artifact path is invalid.")
    );
  }
  let realPath: string;
  try {
    realPath = await fs.realpath(fullPath);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "ArtifactReadFailed",
        "A prepared artifact could not be resolved.",
        { cause: error }
      )
    );
  }
  if (!isWithinProject(root, realPath)) {
    return err(
      new PreparedPublicationError("InvalidArtifactPath", "A prepared artifact path is invalid.")
    );
  }
  const read = await readFileBounded(realPath, SHAREPOINT_PACKAGE_LIMITS.inputBytes);
  if (read.isErr()) {
    return err(read.error);
  }
  if (sha256(read.value) !== artifact.sha256) {
    return err(
      new PreparedPublicationError(
        "ArtifactHashMismatch",
        "A prepared artifact no longer matches its recorded hash."
      )
    );
  }
  return ok(read.value);
}

export async function verifyPreparedSharePointPackage(
  journal: PreparedPublicationJournal,
  bytes: Buffer
): Promise<Result<void, PreparedPublicationError | SharePointPackageInspectionError>> {
  if (sha256(bytes) !== journal.sharePointArtifact.sha256) {
    return err(
      new PreparedPublicationError(
        "SharePointArtifactHashMismatch",
        "The SharePoint package does not match the prepared artifact."
      )
    );
  }
  const identity = await inspectSharePointPackageIdentity(bytes);
  if (identity.isErr()) {
    return err(identity.error);
  }
  if (
    identity.value.productId.toLowerCase() !== journal.sourceSharePoint.productId.toLowerCase() ||
    identity.value.version !== journal.sourceSharePoint.version ||
    identity.value.skipFeatureDeployment !== journal.sourceSharePoint.skipFeatureDeployment
  ) {
    return err(
      new PreparedPublicationError(
        "SharePointArtifactIdentityMismatch",
        "The SharePoint package identity does not match its prepared source."
      )
    );
  }
  return ok(undefined);
}

export function verifyPreparedApplicationPackage(
  journal: PreparedPublicationJournal,
  bytes: Buffer
): Result<void, PreparedPublicationError> {
  if (sha256(bytes) !== journal.applicationArtifact.sha256) {
    return err(
      new PreparedPublicationError(
        "ApplicationArtifactHashMismatch",
        "The application package does not match the prepared artifact."
      )
    );
  }
  const entries = readBoundedArchive(bytes);
  if (entries.isErr()) {
    return err(
      new PreparedPublicationError(
        "InvalidApplicationPackage",
        "The prepared application package is invalid.",
        { cause: entries.error }
      )
    );
  }
  const byName = new Map(entries.value.map((entry) => [entry.name, entry.data]));
  const manifestBytes = byName.get("manifest.json");
  if (manifestBytes === undefined) {
    return err(
      new PreparedPublicationError(
        "InvalidApplicationPackage",
        "The prepared application package is missing manifest.json."
      )
    );
  }
  const manifest = jsonObject(manifestBytes, "manifest.json");
  if (manifest.isErr()) {
    return err(manifest.error);
  }
  if (
    manifest.value["id"] !== journal.applicationArtifact.appId ||
    manifest.value["version"] !== journal.applicationArtifact.version
  ) {
    return err(
      new PreparedPublicationError(
        "ApplicationArtifactIdentityMismatch",
        "The application package identity does not match its prepared source."
      )
    );
  }
  const agentReference = getAgentReference(manifest.value);
  if (agentReference.isErr()) {
    return err(agentReference.error);
  }
  const agentFile = agentReference.value.reference["file"];
  if (typeof agentFile !== "string") {
    return err(
      new PreparedPublicationError(
        "InvalidAgentReference",
        "The application agent reference is invalid."
      )
    );
  }
  const agentBytes = byName.get(agentFile);
  if (agentBytes === undefined) {
    return err(
      new PreparedPublicationError(
        "MissingAgentManifest",
        "The prepared application package is missing its declarative agent manifest."
      )
    );
  }
  const agent = jsonObject(agentBytes, agentFile);
  if (agent.isErr()) {
    return err(agent.error);
  }
  const actionsValue = agent.value["actions"];
  if (!Array.isArray(actionsValue)) {
    return err(
      new PreparedPublicationError(
        "InvalidAgentActions",
        "The prepared declarative agent has no plugin actions."
      )
    );
  }
  const pluginName = `plugins/spfx-api-${journal.plugin.sha256.slice(0, 16)}.json`;
  const pluginEntry = byName.get(pluginName);
  const references = actionsValue.filter(
    (action) => isRecord(action) && action["file"] === pluginName
  );
  if (
    pluginEntry === undefined ||
    references.length !== 1 ||
    sha256(pluginEntry) !== journal.plugin.sha256
  ) {
    return err(
      new PreparedPublicationError(
        "PluginArtifactLineageMismatch",
        "The application package does not contain the prepared API plugin."
      )
    );
  }
  return ok(undefined);
}

export async function verifyPreparedSources(
  root: string,
  journal: PreparedPublicationJournal
): Promise<Result<void, PreparedPublicationError>> {
  const sources = [journal.sourceSharePoint, journal.sourceApplication];
  for (const source of sources) {
    const fullPath = path.resolve(root, ...source.path.split("/"));
    if (!isWithinProject(root, fullPath)) {
      return err(
        new PreparedPublicationError("InvalidSourcePath", "A publication source path is invalid.")
      );
    }
    let realPath: string;
    try {
      realPath = await fs.realpath(fullPath);
    } catch (error) {
      return err(
        new PreparedPublicationError(
          "SourceReadFailed",
          "A publication source could not be resolved.",
          { cause: error }
        )
      );
    }
    if (!isWithinProject(root, realPath)) {
      return err(
        new PreparedPublicationError("InvalidSourcePath", "A publication source path is invalid.")
      );
    }
    const read = await readFileBounded(realPath, SHAREPOINT_PACKAGE_LIMITS.inputBytes);
    if (read.isErr()) {
      return err(read.error);
    }
    if (sha256(read.value) !== source.sha256) {
      return err(
        new PreparedPublicationError(
          "SourceHashMismatch",
          "A publication source changed after preparation."
        )
      );
    }
  }
  return ok(undefined);
}

export function isAmbiguousPublicationState(journal: PreparedPublicationJournal): boolean {
  return (
    journal.state === "sharePointUploadInFlight" ||
    journal.state === "sharePointDeployInFlight" ||
    journal.state === "appSubmissionInFlight" ||
    journal.state === "unknown" ||
    journal.sharePointArtifact.uploadState === "inFlight" ||
    journal.sharePointArtifact.deployState === "inFlight" ||
    journal.applicationArtifact.submissionState === "inFlight" ||
    journal.sharePointArtifact.uploadState === "unknown" ||
    journal.sharePointArtifact.deployState === "unknown" ||
    journal.applicationArtifact.submissionState === "unknown"
  );
}
