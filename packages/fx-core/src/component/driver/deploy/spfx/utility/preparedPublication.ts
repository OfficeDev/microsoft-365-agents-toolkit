// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createHash, randomUUID } from "crypto";
import fs from "fs-extra";
import path from "path";
import properLock from "proper-lockfile";
import {
  ApiPluginManifestConverter,
  AppManifestUtils,
  DeclarativeAgentManifestConverter,
  TeamsManifestConverter,
} from "@microsoft/teamsfx-api";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";
import {
  inspectSharePointPackage,
  inspectSharePointPackageIdentity,
  readBoundedArchive,
  SHAREPOINT_PACKAGE_LIMITS,
  SharePointPackageInspectionError,
} from "./packageFile";
import {
  assembleApplicationPackage,
  SharePointApplicationPackageError,
} from "./applicationPackage";

const journalVersion = 2;
const previousTransformationIdentity = "spfx-copilot-plugin-v2";
const transformationIdentity = "spfx-copilot-plugin-v3";
const publicationDirectory = path.join(".fx", "sharepoint-publications");
export const preparedRequestSettlementTimeoutMs = 30_000;
const tenantMcpUrlPlaceholder = /^\{\{\s*TENANT_MCP_URL\s*\}\}$/;
const placeholderPattern = /\{\{\s*([A-Z][A-Z0-9_]*)\s*\}\}/;
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
export type PreparedPublicationErrorStage =
  | "preparedPublication"
  | "sharePointUpload"
  | "sharePointDeploy"
  | "sharePointVerification"
  | "applicationPreflight"
  | "approval"
  | "applicationSubmission";

export interface PreparedPublicationJournal {
  journalVersion: 2;
  publicationId: string;
  state: PreparedPublicationState;
  identity: {
    tenantId: string;
    tenantOrigin: string;
    catalogUrl: string;
    appId: string;
    productId: string;
  };
  transformation: {
    id: string;
  };
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
    bindingSha256?: string;
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

function isSupportedTransformationIdentity(value: unknown): value is string {
  return value === transformationIdentity || value === previousTransformationIdentity;
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

export function createPreparedRequestSignal(signal?: AbortSignal): AbortSignal {
  const timeoutSignal = AbortSignal.timeout(preparedRequestSettlementTimeoutMs);
  return signal ? AbortSignal.any([signal, timeoutSignal]) : timeoutSignal;
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
  tenantContext: PreparedPublicationTenantContext;
  signal?: AbortSignal;
}

export interface PreparedPublicationTenantContext {
  tenantId: string;
  tenantUrl: string;
  catalogUrl: string;
  productId: string;
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

export function createPreparedPublicationErrorData(
  error: unknown,
  journal: PreparedPublicationJournal | undefined,
  recoveryPath: string | undefined,
  stage: PreparedPublicationErrorStage
): Record<string, unknown> {
  const errorRecord = isRecord(error) ? error : undefined;
  const causeValue = errorRecord?.["cause"] ?? errorRecord?.["innerError"];
  const cause = isRecord(causeValue) ? causeValue : errorRecord;
  const responseValue = cause?.["response"];
  const response = isRecord(responseValue) ? responseValue : undefined;
  const responseDataValue = response?.["data"];
  const responseData = isRecord(responseDataValue) ? responseDataValue : undefined;
  const graphErrorValue = responseData?.["error"];
  const graphError = isRecord(graphErrorValue) ? graphErrorValue : undefined;
  const headersValue = response?.["headers"];
  const headers = isRecord(headersValue) ? headersValue : undefined;
  const serviceCode = graphError?.["code"] ?? responseData?.["code"];
  const requestId = headers?.["request-id"] ?? headers?.["client-request-id"];
  const errorType = errorRecord?.["name"];
  return {
    stage,
    errorType:
      typeof errorType === "string" && /^[A-Za-z][A-Za-z0-9]{0,63}$/.test(errorType)
        ? errorType
        : "Error",
    publicationId: journal?.publicationId,
    lastDurableState: journal?.state,
    lastSharePointVerifiedState:
      journal?.state === "sharePointVerified" ||
      journal?.state === "appSubmissionInFlight" ||
      journal?.state === "unknown" ||
      journal?.state === "rejected" ||
      journal?.state === "submitted"
        ? "sharePointVerified"
        : undefined,
    tenantId: journal?.identity.tenantId,
    appId: journal?.applicationArtifact.appId,
    productId: journal?.sourceSharePoint.productId,
    sourcePackageSha256: journal?.sourceSharePoint.sha256,
    sourceAppPackageSha256: journal?.sourceApplication.sha256,
    pluginSha256: journal?.plugin.sha256,
    sharePointPackageSha256: journal?.sharePointArtifact.sha256,
    appPackageSha256: journal?.applicationArtifact.sha256,
    recoveryPath,
    httpStatus: typeof response?.["status"] === "number" ? response["status"] : undefined,
    serviceCode:
      typeof serviceCode === "string" && /^[A-Za-z0-9._-]{1,100}$/.test(serviceCode)
        ? serviceCode
        : undefined,
    requestId:
      typeof requestId === "string" && /^[A-Za-z0-9._:-]{1,100}$/.test(requestId)
        ? requestId
        : undefined,
  };
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
    const chunks: Buffer[] = [];
    let bytesRead = 0;
    while (bytesRead <= maxBytes) {
      const remaining = maxBytes + 1 - bytesRead;
      const chunk = Buffer.alloc(Math.min(64 * 1024, remaining));
      const readResult = await fs.read(descriptor, chunk, 0, chunk.length, bytesRead);
      if (readResult.bytesRead === 0) {
        break;
      }
      chunks.push(chunk.subarray(0, readResult.bytesRead));
      bytesRead += readResult.bytesRead;
      if (bytesRead > maxBytes) {
        return err(
          new PreparedPublicationError(
            "InputTooLarge",
            "A publication input exceeds safe size limits."
          )
        );
      }
    }
    const finalStats = await fs.fstat(descriptor);
    // ctime is excluded: Windows updates it lazily after writes without any content change.
    if (finalStats.size !== bytesRead || finalStats.mtimeMs !== stats.mtimeMs) {
      return err(
        new PreparedPublicationError(
          "InputChangedDuringRead",
          "A publication input changed while it was read."
        )
      );
    }
    const bytes = Buffer.concat(chunks, bytesRead);
    await fs.close(descriptor);
    descriptor = undefined;
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
  identity: PreparedPublicationTenantContext & {
    appId: string;
    productId: string;
    sourceSharePointSha256: string;
    sourceApplicationSha256: string;
  }
): Promise<Result<PreparedPublicationJournal | undefined, PreparedPublicationError>> {
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
    const journalBytes = await readFileBounded(
      path.resolve(root, journalPath),
      SHAREPOINT_PACKAGE_LIMITS.metadataBytes
    );
    if (journalBytes.isErr()) {
      return err(
        new PreparedPublicationError(
          "PublicationRecoveryCheckFailed",
          "An existing publication record could not be checked safely.",
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
          "PublicationRecoveryCheckFailed",
          "An existing publication record could not be checked safely.",
          { cause: error }
        )
      );
    }
    if (isRecord(raw) && raw["journalVersion"] === 1 && raw["state"] !== "submitted") {
      return err(
        new PreparedPublicationError(
          "LegacyPreparedPublicationRequiresReconciliation",
          "An active legacy publication record lacks tenant-bound lineage and must be reconciled before new prepared work."
        )
      );
    }
    if (isRecord(raw) && raw["journalVersion"] === 1 && raw["state"] === "submitted") {
      continue;
    }
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
      journal.identity.tenantId.toLowerCase() === identity.tenantId.toLowerCase() &&
      journal.identity.appId.toLowerCase() === identity.appId.toLowerCase() &&
      journal.identity.productId.toLowerCase() === identity.productId.toLowerCase() &&
      journal.identity.tenantOrigin.toLowerCase() ===
        new URL(identity.tenantUrl).origin.toLowerCase() &&
      journal.identity.catalogUrl.toLowerCase() === identity.catalogUrl.toLowerCase()
    ) {
      if (
        ["prepared", "sharePointUploaded", "sharePointVerified", "submitted"].includes(
          journal.state
        ) &&
        journal.transformation.id === transformationIdentity &&
        journal.sourceSharePoint.sha256 === identity.sourceSharePointSha256 &&
        journal.sourceApplication.sha256 === identity.sourceApplicationSha256
      ) {
        return ok(journal);
      }
      if (journal.state === "submitted") {
        continue;
      }
      return err(
        new PreparedPublicationError(
          "UnresolvedPreparedPublication",
          "A prepared publication for this tenant, application, and SharePoint solution is still active, rejected, or uncertain and requires reconciliation before a new lineage can be created."
        )
      );
    }
  }
  return ok(undefined);
}

function replaceTenantValues(
  pluginBytes: Buffer,
  productId: string,
  tenantContext: PreparedPublicationTenantContext
): Result<Buffer, PreparedPublicationError> {
  let tenantUrl: URL;
  let catalogUrl: URL;
  try {
    tenantUrl = new URL(tenantContext.tenantUrl);
    catalogUrl = new URL(tenantContext.catalogUrl);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InvalidPreparedTenantContext",
        "The selected tenant or app catalog URL is invalid.",
        { cause: error }
      )
    );
  }
  if (
    !guidPattern.test(tenantContext.tenantId) ||
    tenantUrl.protocol !== "https:" ||
    tenantUrl.username ||
    tenantUrl.password ||
    tenantUrl.search ||
    tenantUrl.hash ||
    catalogUrl.protocol !== "https:" ||
    catalogUrl.username ||
    catalogUrl.password ||
    catalogUrl.search ||
    catalogUrl.hash ||
    tenantUrl.origin !== catalogUrl.origin ||
    productId.toLowerCase() !== tenantContext.productId.toLowerCase()
  ) {
    return err(
      new PreparedPublicationError(
        "InvalidPreparedTenantContext",
        "The selected tenant, app catalog, and SharePoint solution identity do not match."
      )
    );
  }
  let plugin: Record<string, unknown>;
  try {
    const parsed: unknown = JSON.parse(pluginBytes.toString("utf8"));
    if (!isRecord(parsed)) {
      throw new Error("The API plugin root must be an object.");
    }
    plugin = parsed;
  } catch (error) {
    return err(
      new PreparedPublicationError("InvalidPluginManifest", "The embedded API plugin is invalid.", {
        cause: error,
      })
    );
  }
  const runtimes = plugin["runtimes"];
  if (!Array.isArray(runtimes) || !runtimes.every(isRecord)) {
    return err(
      new PreparedPublicationError(
        "InvalidPluginRuntimes",
        "The embedded API plugin runtimes are invalid."
      )
    );
  }
  const endpoint = new URL("/_api/mcp/spfx", catalogUrl.origin);
  endpoint.searchParams.set("solutionId", productId);
  let resolvedEndpointCount = 0;
  for (const runtime of runtimes) {
    const spec = runtime["spec"];
    if (runtime["type"] !== "RemoteMCPServer" || !isRecord(spec)) {
      continue;
    }
    const url = spec["url"];
    if (typeof url !== "string") {
      continue;
    }
    if (tenantMcpUrlPlaceholder.test(url.trim())) {
      spec["url"] = endpoint.toString();
      resolvedEndpointCount += 1;
    } else if (placeholderPattern.test(url)) {
      return err(
        new PreparedPublicationError(
          "UnresolvedPluginPlaceholder",
          "A registered API endpoint contains an unsupported tenant placeholder."
        )
      );
    }
    for (const [key, value] of Object.entries(spec)) {
      if (
        /csp|content.?security/i.test(key) &&
        JSON.stringify(value).includes("{{TENANT_ORIGIN}}")
      ) {
        return err(
          new PreparedPublicationError(
            "UnqualifiedTenantOriginField",
            "The tenant-origin CSP field is not qualified for this publication format."
          )
        );
      }
    }
  }
  if (resolvedEndpointCount !== 1) {
    return err(
      new PreparedPublicationError(
        "AmbiguousPluginEndpoint",
        "The embedded API plugin must register exactly one tenant-bound SharePoint endpoint."
      )
    );
  }
  return ok(Buffer.from(JSON.stringify(plugin, null, 2), "utf8"));
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
  | {
      container: Record<string, unknown>;
      property: "declarativeAgents" | "declarativeCopilots";
      reference: Record<string, unknown>;
    }
  | undefined,
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
  if (references.length > 1) {
    return err(
      new PreparedPublicationError(
        "UnsupportedApplicationPackage",
        "The application package contains multiple declarative agent references."
      )
    );
  }
  return ok(references[0]);
}

function prepareApplicationPackage(
  bytes: Buffer,
  pluginBytes: Buffer,
  sourcePluginBytes: Buffer,
  productId: string
): Result<
  { bytes: Buffer; appId: string; version: string; pluginBindingSha256: string },
  PreparedPublicationError
> {
  const assembled = assembleApplicationPackage(bytes, pluginBytes, sourcePluginBytes, productId);
  if (assembled.isErr()) {
    const error: SharePointApplicationPackageError = assembled.error;
    return err(
      new PreparedPublicationError(error.code, error.message, {
        cause: error,
      })
    );
  }
  return ok({
    bytes: assembled.value.bytes,
    appId: assembled.value.appId,
    version: assembled.value.appVersion,
    pluginBindingSha256: assembled.value.pluginBindingSha256,
  });
}

function getTeamsSchemaUrl(version: unknown): string | undefined {
  if (typeof version !== "string") {
    return undefined;
  }
  if (["devPreview", "m365DevPreview", "m365devpreview"].includes(version)) {
    return "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json";
  }
  if (/^1\.(?:[0-9]|1[0-9]|2[0-9]|30)$/.test(version)) {
    return `https://developer.microsoft.com/json-schemas/teams/v${version}/MicrosoftTeams.schema.json`;
  }
  return undefined;
}

function getPluginSchemaUrl(version: unknown): string | undefined {
  if (typeof version !== "string" || !/^v2\.[1-4]$/.test(version)) {
    return undefined;
  }
  return `https://developer.microsoft.com/json-schemas/copilot/plugin/${version}/schema.json`;
}

function getAgentSchemaUrl(version: unknown): string | undefined {
  if (
    typeof version !== "string" ||
    !["v1.0", "v1.2", "v1.3", "v1.4", "v1.5", "v1.6", "v1.7", "v1.8", "v1.9"].includes(version)
  ) {
    return undefined;
  }
  return `https://developer.microsoft.com/json-schemas/copilot/declarative-agent/${version}/schema.json`;
}

async function validatePreparedApplicationPackage(
  bytes: Buffer
): Promise<Result<void, PreparedPublicationError>> {
  const entries = readBoundedArchive(bytes);
  if (entries.isErr()) {
    return err(entries.error);
  }
  const byName = new Map(entries.value.map((entry) => [entry.name, entry.data]));
  const manifestBytes = byName.get("manifest.json");
  if (!manifestBytes) {
    return err(
      new PreparedPublicationError(
        "ApplicationSchemaValidationFailed",
        "The staged application manifest is missing."
      )
    );
  }
  const manifestResult = jsonObject(manifestBytes, "manifest.json");
  if (manifestResult.isErr()) {
    return err(manifestResult.error);
  }
  const manifest = manifestResult.value;
  const teamsSchemaUrl = getTeamsSchemaUrl(manifest["manifestVersion"]);
  if (!teamsSchemaUrl) {
    return err(
      new PreparedPublicationError(
        "UnsupportedTeamsManifestVersion",
        "The staged application manifest version is not supported."
      )
    );
  }
  try {
    const schema = await AppManifestUtils.fetchSchema(teamsSchemaUrl);
    const typedManifest = TeamsManifestConverter.jsonToManifest(manifestBytes.toString("utf8"));
    const schemaErrors = await AppManifestUtils.validateAgainstSchema(typedManifest, schema);
    if (schemaErrors.length > 0) {
      return err(
        new PreparedPublicationError(
          "ApplicationSchemaValidationFailed",
          "The staged application manifest does not satisfy its supported schema."
        )
      );
    }

    const validatePlugin = async (pluginPath: string): Promise<boolean> => {
      if (!isSafeRelativePath(pluginPath)) {
        return false;
      }
      const pluginBytes = byName.get(pluginPath);
      if (!pluginBytes) {
        return false;
      }
      const pluginObject = jsonObject(pluginBytes, pluginPath);
      if (pluginObject.isErr()) {
        return false;
      }
      const pluginSchemaUrl = getPluginSchemaUrl(pluginObject.value["schema_version"]);
      if (!pluginSchemaUrl) {
        return false;
      }
      const pluginSchema = await AppManifestUtils.fetchSchema(pluginSchemaUrl);
      const typedPlugin = ApiPluginManifestConverter.jsonToManifest(pluginBytes.toString("utf8"));
      return (await AppManifestUtils.validateAgainstSchema(typedPlugin, pluginSchema)).length === 0;
    };

    const agentReferences: string[] = [];
    for (const containerName of ["copilotAgents", "copilotExtensions"]) {
      const container = manifest[containerName];
      if (!isRecord(container)) {
        continue;
      }
      const references =
        container[containerName === "copilotAgents" ? "declarativeAgents" : "declarativeCopilots"];
      if (!Array.isArray(references)) {
        continue;
      }
      for (const reference of references) {
        if (isRecord(reference) && typeof reference["file"] === "string") {
          agentReferences.push(reference["file"]);
        }
      }
    }
    const pluginReferences = new Set<string>();
    for (const agentConnector of Array.isArray(manifest["agentConnectors"])
      ? manifest["agentConnectors"]
      : []) {
      if (
        isRecord(agentConnector) &&
        isRecord(agentConnector["toolSource"]) &&
        isRecord(agentConnector["toolSource"]["plugin"]) &&
        typeof agentConnector["toolSource"]["plugin"]["file"] === "string"
      ) {
        pluginReferences.add(agentConnector["toolSource"]["plugin"]["file"]);
      }
    }
    for (const agentPath of agentReferences) {
      if (!isSafeRelativePath(agentPath)) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "The staged application contains an invalid agent reference."
          )
        );
      }
      const agentBytes = byName.get(agentPath);
      if (!agentBytes) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "The staged application references a missing agent manifest."
          )
        );
      }
      const agentObject = jsonObject(agentBytes, agentPath);
      if (agentObject.isErr()) {
        return err(agentObject.error);
      }
      const agentSchemaUrl = getAgentSchemaUrl(agentObject.value["version"]);
      if (!agentSchemaUrl) {
        return err(
          new PreparedPublicationError(
            "UnsupportedAgentManifestVersion",
            "The staged declarative agent version is not supported."
          )
        );
      }
      const agentSchema = await AppManifestUtils.fetchSchema(agentSchemaUrl);
      const typedAgent = DeclarativeAgentManifestConverter.jsonToManifest(
        agentBytes.toString("utf8")
      );
      if ((await AppManifestUtils.validateAgainstSchema(typedAgent, agentSchema)).length > 0) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "A staged declarative agent does not satisfy its supported schema."
          )
        );
      }
      const actions = agentObject.value["actions"];
      if (!Array.isArray(actions)) {
        continue;
      }
      for (const action of actions) {
        if (isRecord(action) && typeof action["file"] === "string") {
          pluginReferences.add(action["file"]);
        }
      }
    }
    for (const pluginPath of pluginReferences) {
      if (!(await validatePlugin(pluginPath))) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "A staged API plugin reference is missing or does not satisfy its supported schema."
          )
        );
      }
    }
    return ok(undefined);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "ApplicationSchemaValidationFailed",
        "The staged application package could not be validated against its supported schemas.",
        { cause: error }
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

function getIdentityLockPath(
  root: string,
  identity: { tenantId: string; appId: string; productId: string }
): string {
  const key = sha256(
    Buffer.from(
      `${identity.tenantId.toLowerCase()}\n${identity.appId.toLowerCase()}\n${identity.productId.toLowerCase()}`,
      "utf8"
    )
  );
  return path.join(root, publicationDirectory, `_identity-${key}`);
}

async function withIdentityLock<T>(
  root: string,
  identity: { tenantId: string; appId: string; productId: string },
  operation: () => Promise<T>
): Promise<T> {
  let release: () => Promise<void>;
  try {
    const lockDirectory = getIdentityLockPath(root, identity);
    await fs.ensureDir(lockDirectory);
    release = await properLock.lock(lockDirectory, { retries: 0 });
  } catch (error) {
    throw new PreparedPublicationError(
      "PublicationIdentityLockUnavailable",
      "A conflicting publication for this tenant, application, or SharePoint solution is active.",
      { cause: error }
    );
  }
  let result: T;
  try {
    result = await operation();
  } catch (error) {
    await release().catch(() => undefined);
    throw error;
  }
  try {
    await release();
  } catch (error) {
    throw new PreparedPublicationError(
      "PublicationIdentityLockReleaseFailed",
      "The publication identity lock could not be released.",
      { cause: error }
    );
  }
  return result;
}

export async function withPreparedPublicationIdentityLock<T>(
  root: string,
  journal: PreparedPublicationJournal,
  operation: () => Promise<T>
): Promise<T> {
  if (journal.journalVersion !== journalVersion) {
    throw new PreparedPublicationError(
      "LegacyPreparedPublicationRequiresReconciliation",
      "This publication journal does not contain the current tenant-bound lineage."
    );
  }
  return withIdentityLock(root, journal.identity, operation);
}

export async function getPreparedSharePointProductId(
  projectPath: string,
  packagePath: string
): Promise<Result<string, PreparedPublicationError | SharePointPackageInspectionError>> {
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
  const sourcePackage = await readProjectInput(root, packagePath);
  if (sourcePackage.isErr()) {
    return err(sourcePackage.error);
  }
  const inspection = await inspectSharePointPackage(sourcePackage.value.bytes);
  return inspection.isErr() ? err(inspection.error) : ok(inspection.value.productId);
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
  if (input.signal?.aborted) {
    return err(
      new PreparedPublicationError(
        "PublicationCancelled",
        "Prepared publication was cancelled before tenant discovery."
      )
    );
  }
  const sharePointPackage = await inspectSharePointPackage(sourcePackage.value.bytes);
  if (sharePointPackage.isErr()) {
    return err(sharePointPackage.error);
  }
  if (
    sharePointPackage.value.productId.toLowerCase() !== input.tenantContext.productId.toLowerCase()
  ) {
    return err(
      new PreparedPublicationError(
        "SelectedProductMismatch",
        "The selected tenant context does not match the inspected SharePoint solution."
      )
    );
  }
  const resolvedPlugin = replaceTenantValues(
    sharePointPackage.value.pluginManifest,
    sharePointPackage.value.productId,
    input.tenantContext
  );
  if (resolvedPlugin.isErr()) {
    return err(resolvedPlugin.error);
  }
  const pluginSha256 = sha256(resolvedPlugin.value);
  const appPackage = prepareApplicationPackage(
    sourceAppPackage.value.bytes,
    resolvedPlugin.value,
    sharePointPackage.value.pluginManifest,
    sharePointPackage.value.productId
  );
  if (appPackage.isErr()) {
    return err(appPackage.error);
  }
  const validatedApplication = await validatePreparedApplicationPackage(appPackage.value.bytes);
  if (validatedApplication.isErr()) {
    return err(validatedApplication.error);
  }
  let tenantUrl: URL;
  try {
    tenantUrl = new URL(input.tenantContext.tenantUrl);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InvalidPreparedTenantContext",
        "The selected tenant URL is invalid.",
        { cause: error }
      )
    );
  }
  const identity = {
    tenantId: input.tenantContext.tenantId,
    tenantUrl: input.tenantContext.tenantUrl,
    catalogUrl: input.tenantContext.catalogUrl,
    appId: appPackage.value.appId,
    productId: sharePointPackage.value.productId,
  };
  try {
    return await withIdentityLock(root, identity, async () => {
      if (input.signal?.aborted) {
        return err(
          new PreparedPublicationError(
            "PublicationCancelled",
            "Prepared publication was cancelled before staging."
          )
        );
      }
      const recoveryCheck = await findUnresolvedPublication(root, {
        ...input.tenantContext,
        appId: appPackage.value.appId,
        productId: sharePointPackage.value.productId,
        sourceSharePointSha256: sha256(sourcePackage.value.bytes),
        sourceApplicationSha256: sha256(sourceAppPackage.value.bytes),
      });
      if (recoveryCheck.isErr()) {
        return err(recoveryCheck.error);
      }
      if (recoveryCheck.value) {
        const journal = recoveryCheck.value;
        const sources = await verifyPreparedSources(root, journal);
        if (sources.isErr()) {
          return err(sources.error);
        }
        const sharePointArtifact = await verifyPreparedArtifact(root, journal.sharePointArtifact);
        if (sharePointArtifact.isErr()) {
          return err(sharePointArtifact.error);
        }
        const sharePointLineage = await verifyPreparedSharePointPackage(
          journal,
          sharePointArtifact.value
        );
        if (sharePointLineage.isErr()) {
          return err(sharePointLineage.error);
        }
        const appArtifact = await verifyPreparedArtifact(root, journal.applicationArtifact);
        if (appArtifact.isErr()) {
          return err(appArtifact.error);
        }
        const appLineage = verifyPreparedApplicationPackage(journal, appArtifact.value);
        if (appLineage.isErr()) {
          return err(appLineage.error);
        }
        return ok({
          publicationId: journal.publicationId,
          preparedPublicationPath: projectRelative(
            root,
            path.join(root, publicationDirectory, journal.publicationId, "journal.json")
          ),
          sourcePackageSha256: journal.sourceSharePoint.sha256,
          sourceAppPackageSha256: journal.sourceApplication.sha256,
          pluginSha256: journal.plugin.sha256,
          sharePointPackagePath: journal.sharePointArtifact.path,
          sharePointPackageSha256: journal.sharePointArtifact.sha256,
          appPackagePath: journal.applicationArtifact.path,
          appPackageSha256: journal.applicationArtifact.sha256,
        });
      }
      const publicationId = randomUUID();
      const stagingDirectory = path.resolve(root, publicationDirectory, publicationId);
      if (!isWithinProject(root, stagingDirectory)) {
        return err(
          new PreparedPublicationError(
            "InvalidStagingPath",
            "The publication staging path is invalid."
          )
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
        if (input.signal?.aborted) {
          await fs.remove(stagingDirectory);
          return err(
            new PreparedPublicationError(
              "PublicationCancelled",
              "Prepared publication was cancelled before the journal was committed."
            )
          );
        }
        await atomicWrite(applicationArtifactPath, appPackage.value.bytes);
        const journal: PreparedPublicationJournal = {
          journalVersion,
          publicationId,
          state: "prepared",
          identity: {
            tenantId: input.tenantContext.tenantId,
            tenantOrigin: tenantUrl.origin,
            catalogUrl: input.tenantContext.catalogUrl,
            appId: appPackage.value.appId,
            productId: sharePointPackage.value.productId,
          },
          transformation: { id: transformationIdentity },
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
          plugin: {
            sha256: pluginSha256,
            bindingSha256: appPackage.value.pluginBindingSha256,
          },
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
                { cause: error }
              )
        );
      }
    });
  } catch (error) {
    return err(
      error instanceof PreparedPublicationError
        ? error
        : new PreparedPublicationError(
            "PublicationPrepareFailed",
            "The publication artifacts could not be prepared.",
            { cause: error }
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
    !isRecord(value["identity"]) ||
    typeof value["identity"]["tenantId"] !== "string" ||
    !guidPattern.test(value["identity"]["tenantId"]) ||
    typeof value["identity"]["tenantOrigin"] !== "string" ||
    typeof value["identity"]["catalogUrl"] !== "string" ||
    typeof value["identity"]["appId"] !== "string" ||
    typeof value["identity"]["productId"] !== "string" ||
    !isRecord(value["transformation"]) ||
    !isSupportedTransformationIdentity(value["transformation"]["id"]) ||
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
    (value["plugin"]["bindingSha256"] !== undefined &&
      (typeof value["plugin"]["bindingSha256"] !== "string" ||
        !/^[0-9a-f]{64}$/i.test(value["plugin"]["bindingSha256"]))) ||
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
  const identity = value["identity"];
  const transformation = value["transformation"];
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
  const pluginBindingSha256 = plugin["bindingSha256"];
  const sharePointArtifactSha256 = sharePointArtifact["sha256"];
  const sharePointUploadState = sharePointArtifact["uploadState"];
  const sharePointDeployState = sharePointArtifact["deployState"];
  const applicationArtifactSha256 = applicationArtifact["sha256"];
  const applicationArtifactAppId = applicationArtifact["appId"];
  const applicationArtifactVersion = applicationArtifact["version"];
  const applicationSubmissionState = applicationArtifact["submissionState"];
  const tenantId = identity["tenantId"];
  const tenantOrigin = identity["tenantOrigin"];
  const catalogUrl = identity["catalogUrl"];
  const identityAppId = identity["appId"];
  const identityProductId = identity["productId"];
  const transformationId = transformation["id"];
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
    (pluginBindingSha256 !== undefined && typeof pluginBindingSha256 !== "string") ||
    typeof sharePointArtifactSha256 !== "string" ||
    !isWriteState(sharePointUploadState) ||
    !isWriteState(sharePointDeployState) ||
    typeof applicationArtifactSha256 !== "string" ||
    typeof applicationArtifactAppId !== "string" ||
    typeof applicationArtifactVersion !== "string" ||
    typeof tenantId !== "string" ||
    typeof tenantOrigin !== "string" ||
    typeof catalogUrl !== "string" ||
    typeof identityAppId !== "string" ||
    typeof identityProductId !== "string" ||
    typeof transformationId !== "string" ||
    identityAppId.toLowerCase() !== sourceApplicationAppId.toLowerCase() ||
    identityProductId.toLowerCase() !== sourceSharePointProductId.toLowerCase() ||
    (transformationId !== transformationIdentity &&
      transformationId !== previousTransformationIdentity) ||
    !isPreparedCatalogUrl(catalogUrl, tenantOrigin) ||
    !isWriteState(applicationSubmissionState)
  ) {
    return undefined;
  }
  return {
    journalVersion: 2,
    publicationId: value["publicationId"],
    state: value["state"],
    identity: {
      tenantId,
      tenantOrigin,
      catalogUrl,
      appId: identityAppId,
      productId: identityProductId,
    },
    transformation: { id: transformationId },
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
    plugin: {
      sha256: pluginSha256,
      ...(typeof pluginBindingSha256 === "string" ? { bindingSha256: pluginBindingSha256 } : {}),
    },
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
  const pluginBindingSha256 = journal.plugin.bindingSha256 ?? journal.plugin.sha256;
  const pluginName = `plugins/spfx-api-${pluginBindingSha256.slice(0, 16)}.json`;
  const pluginId = `spfx-api-${pluginBindingSha256.slice(0, 16)}`;
  const pluginEntry = byName.get(pluginName);
  let referenceCount = 0;
  if (agentReference.value) {
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
    referenceCount = actionsValue.filter(
      (action) => isRecord(action) && action["id"] === pluginId && action["file"] === pluginName
    ).length;
  } else {
    const connectors = manifest.value["agentConnectors"];
    if (!Array.isArray(connectors)) {
      return err(
        new PreparedPublicationError(
          "MissingAgentConnector",
          "The prepared application package is missing its API plugin connector."
        )
      );
    }
    referenceCount = connectors.filter(
      (connector) =>
        isRecord(connector) &&
        connector["id"] === pluginId &&
        isRecord(connector["toolSource"]) &&
        isRecord(connector["toolSource"]["plugin"]) &&
        connector["toolSource"]["plugin"]["file"] === pluginName
    ).length;
  }
  if (
    pluginEntry === undefined ||
    referenceCount !== 1 ||
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
