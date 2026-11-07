// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { createHash } from "crypto";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";
import path from "path";
import { readBoundedArchive, SharePointPackageInspectionError } from "./packageFile";

const managedPluginPrefix = "spfx-api-";
const managedPluginPathPrefix = "plugins/spfx-api-";
const placeholderEndpoint = /^\{\{\s*TENANT_MCP_URL\s*\}\}$/;

export class SharePointApplicationPackageError extends Error {
  public constructor(
    public readonly code: string,
    message: string,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = "SharePointApplicationPackageError";
  }
}

export interface AssembledApplicationPackage {
  bytes: Buffer;
  appId: string;
  appVersion: string;
  pluginSha256: string;
  pluginBindingSha256: string;
  managedReference: {
    kind: "declarativeAgentAction" | "agentConnector";
    id: string;
    path: string;
  };
}

interface AgentReference {
  file: string;
  reference: Record<string, unknown>;
  container: Record<string, unknown>;
  property: "declarativeAgents" | "declarativeCopilots";
}

interface ManagedReferenceCandidate {
  kind: "declarativeAgentAction" | "agentConnector";
  id: unknown;
  path: unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function sha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function packageError(
  code: string,
  message: string,
  cause?: unknown
): SharePointApplicationPackageError {
  return new SharePointApplicationPackageError(code, message, { cause });
}

function parseJsonObject(
  bytes: Buffer,
  name: string
): Result<Record<string, unknown>, SharePointApplicationPackageError> {
  try {
    const parsed: unknown = JSON.parse(bytes.toString("utf8").replace(/^\uFEFF/, ""));
    if (!isRecord(parsed)) {
      return err(
        packageError("InvalidPackageMetadata", `Package metadata "${name}" must be a JSON object.`)
      );
    }
    return ok(parsed);
  } catch (error) {
    return err(
      packageError("InvalidPackageMetadata", `Package metadata "${name}" is invalid.`, error)
    );
  }
}

function getAgentReference(
  manifest: Record<string, unknown>
): Result<AgentReference | undefined, SharePointApplicationPackageError> {
  const references: AgentReference[] = [];
  const referenceTypes: {
    containerName: string;
    property: "declarativeAgents" | "declarativeCopilots";
  }[] = [
    { containerName: "copilotAgents", property: "declarativeAgents" },
    { containerName: "copilotExtensions", property: "declarativeCopilots" },
  ];
  for (const referenceType of referenceTypes) {
    const containerValue = manifest[referenceType.containerName];
    if (containerValue === undefined) {
      continue;
    }
    if (!isRecord(containerValue)) {
      return err(
        packageError(
          "InvalidApplicationPackage",
          "The application declarative-agent container is invalid."
        )
      );
    }
    const values = containerValue[referenceType.property];
    if (values === undefined) {
      continue;
    }
    if (!Array.isArray(values)) {
      return err(
        packageError(
          "InvalidApplicationPackage",
          "The application declarative-agent references are invalid."
        )
      );
    }
    for (const value of values) {
      if (isRecord(value)) {
        references.push({
          container: containerValue,
          property: referenceType.property,
          reference: value,
          file: typeof value["file"] === "string" ? value["file"] : "",
        });
      } else {
        return err(
          packageError(
            "InvalidApplicationPackage",
            "The application declarative-agent reference is invalid."
          )
        );
      }
    }
  }
  if (references.length > 1) {
    return err(
      packageError(
        "UnsupportedApplicationPackage",
        "The application package contains multiple declarative agent references."
      )
    );
  }
  return ok(references[0]);
}

function isSafeRelativePath(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length > 0 &&
    !value.includes("\\") &&
    !value.includes("\0") &&
    !value.startsWith("/") &&
    !/^[A-Za-z]:/.test(value) &&
    path.posix.normalize(value) === value
  );
}

function stableJson(value: unknown): string {
  if (Array.isArray(value)) {
    return `[${value.map(stableJson).join(",")}]`;
  }
  if (isRecord(value)) {
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableJson(value[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value) ?? "null";
}

function normalizeManagedEndpoint(plugin: Record<string, unknown>, productId: string): void {
  const runtimes = plugin["runtimes"];
  if (!Array.isArray(runtimes)) {
    return;
  }
  for (const runtime of runtimes) {
    if (!isRecord(runtime) || runtime["type"] !== "RemoteMCPServer") {
      continue;
    }
    const spec = runtime["spec"];
    if (!isRecord(spec) || typeof spec["url"] !== "string") {
      continue;
    }
    const value = spec["url"].trim();
    if (placeholderEndpoint.test(value)) {
      spec["url"] = "{{TENANT_MCP_URL}}";
      continue;
    }
    try {
      const url = new URL(value);
      if (
        url.protocol === "https:" &&
        !url.username &&
        !url.password &&
        !url.hash &&
        url.pathname === "/_api/mcp/spfx" &&
        url.searchParams.getAll("solutionId").length === 1 &&
        url.searchParams.get("solutionId")?.toLowerCase() === productId.toLowerCase() &&
        [...url.searchParams.keys()].length === 1
      ) {
        spec["url"] = "{{TENANT_MCP_URL}}";
      }
    } catch {
      return;
    }
  }
}

function matchesSourcePlugin(
  candidateBytes: Buffer,
  sourcePlugin: Record<string, unknown>,
  productId: string
): boolean {
  const candidate = parseJsonObject(candidateBytes, "managed API plugin");
  if (candidate.isErr()) {
    return false;
  }
  normalizeManagedEndpoint(candidate.value, productId);
  const normalizedSource = structuredClone(sourcePlugin);
  normalizeManagedEndpoint(normalizedSource, productId);
  return stableJson(candidate.value) === stableJson(normalizedSource);
}

function validateSourcePluginEndpoint(
  plugin: Record<string, unknown>
): Result<void, SharePointApplicationPackageError> {
  const runtimes = plugin["runtimes"];
  if (!Array.isArray(runtimes) || runtimes.length !== 1 || !isRecord(runtimes[0])) {
    return err(
      packageError(
        "UnsupportedPluginRuntimes",
        "The source API plugin must contain one SharePoint Remote MCP runtime."
      )
    );
  }
  const runtime = runtimes[0];
  const spec = runtime["spec"];
  const url = isRecord(spec) ? spec["url"] : undefined;
  if (
    runtime["type"] !== "RemoteMCPServer" ||
    typeof url !== "string" ||
    !placeholderEndpoint.test(url.trim())
  ) {
    return err(
      packageError(
        "InvalidTenantNeutralPlugin",
        "The source API plugin must use the generated tenant endpoint placeholder."
      )
    );
  }
  return ok(undefined);
}

function managedCandidates(
  actions: Record<string, unknown>[],
  connectors: Record<string, unknown>[]
): ManagedReferenceCandidate[] {
  const candidates: ManagedReferenceCandidate[] = [];
  for (const action of actions) {
    const id = action["id"];
    const file = action["file"];
    if (
      (typeof id === "string" && id.startsWith(managedPluginPrefix)) ||
      (typeof file === "string" && file.startsWith(managedPluginPathPrefix))
    ) {
      candidates.push({ kind: "declarativeAgentAction", id, path: file });
    }
  }
  for (const connector of connectors) {
    const id = connector["id"];
    const toolSource = connector["toolSource"];
    const plugin = isRecord(toolSource) ? toolSource["plugin"] : undefined;
    const file = isRecord(plugin) ? plugin["file"] : undefined;
    if (
      (typeof id === "string" && id.startsWith(managedPluginPrefix)) ||
      (typeof file === "string" && file.startsWith(managedPluginPathPrefix))
    ) {
      candidates.push({ kind: "agentConnector", id, path: file });
    }
  }
  return candidates;
}

function isLegacyManagedReference(candidate: ManagedReferenceCandidate, bytes: Buffer): boolean {
  const contentHash = sha256(bytes).slice(0, 16);
  return (
    candidate.id === `${managedPluginPrefix}${contentHash}` &&
    candidate.path === `${managedPluginPathPrefix}${contentHash}.json`
  );
}

function readReferenceArrays(
  manifest: Record<string, unknown>,
  agentManifest: Record<string, unknown> | undefined
): Result<
  { actions: Record<string, unknown>[]; connectors: Record<string, unknown>[] },
  SharePointApplicationPackageError
> {
  let actions: Record<string, unknown>[] = [];
  if (agentManifest) {
    const actionsValue = agentManifest["actions"];
    if (actionsValue !== undefined && !Array.isArray(actionsValue)) {
      return err(
        packageError("InvalidAgentActions", "The declarative agent actions must be an array.")
      );
    }
    if (Array.isArray(actionsValue)) {
      if (!actionsValue.every(isRecord)) {
        return err(
          packageError(
            "InvalidAgentActions",
            "The declarative agent actions contain an invalid entry."
          )
        );
      }
      actions = actionsValue;
    }
  }
  const connectorsValue = manifest["agentConnectors"];
  if (connectorsValue !== undefined && !Array.isArray(connectorsValue)) {
    return err(
      packageError("InvalidAgentConnectors", "The application agent connectors are invalid.")
    );
  }
  const connectors = Array.isArray(connectorsValue) ? connectorsValue : [];
  if (!connectors.every(isRecord)) {
    return err(
      packageError(
        "InvalidAgentConnectors",
        "The application agent connectors contain an invalid entry."
      )
    );
  }
  return ok({ actions, connectors });
}

export function assembleApplicationPackage(
  applicationPackageBytes: Buffer,
  pluginBytes: Buffer,
  sourcePluginBytes: Buffer,
  productId: string
): Result<AssembledApplicationPackage, SharePointApplicationPackageError> {
  const entries = readBoundedArchive(applicationPackageBytes);
  if (entries.isErr()) {
    return err(packageError(entries.error.code, entries.error.message, entries.error));
  }
  const byName = new Map(entries.value.map((entry) => [entry.name, entry.data]));
  const manifestBytes = byName.get("manifest.json");
  if (manifestBytes === undefined) {
    return err(
      packageError("InvalidApplicationPackage", "The application package is missing manifest.json.")
    );
  }
  const manifestResult = parseJsonObject(manifestBytes, "manifest.json");
  if (manifestResult.isErr()) {
    return err(manifestResult.error);
  }
  const manifest = manifestResult.value;
  const appId = manifest["id"];
  const appVersion = manifest["version"];
  if (
    typeof appId !== "string" ||
    !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(appId) ||
    typeof appVersion !== "string" ||
    !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(appVersion)
  ) {
    return err(
      packageError("InvalidApplicationIdentity", "The application identity or version is invalid.")
    );
  }
  const pluginManifestResult = parseJsonObject(pluginBytes, "prepared API plugin");
  if (pluginManifestResult.isErr()) {
    return err(pluginManifestResult.error);
  }
  const sourcePluginResult = parseJsonObject(sourcePluginBytes, "source API plugin");
  if (sourcePluginResult.isErr()) {
    return err(sourcePluginResult.error);
  }
  const pluginManifest = pluginManifestResult.value;
  const sourcePluginManifest = sourcePluginResult.value;
  const sourceEndpoint = validateSourcePluginEndpoint(sourcePluginManifest);
  if (sourceEndpoint.isErr()) {
    return err(sourceEndpoint.error);
  }
  const pluginNamespace = pluginManifest["namespace"];
  const pluginName = pluginManifest["name_for_human"];
  if (
    typeof pluginNamespace !== "string" ||
    !/^[A-Za-z0-9-]+$/.test(pluginNamespace) ||
    typeof pluginName !== "string" ||
    pluginName.trim().length === 0
  ) {
    return err(
      packageError("InvalidPluginIdentity", "The extracted API plugin identity is invalid.")
    );
  }
  const agentReferenceResult = getAgentReference(manifest);
  if (agentReferenceResult.isErr()) {
    return err(agentReferenceResult.error);
  }
  const agentReference = agentReferenceResult.value;
  let agentManifest: Record<string, unknown> | undefined;
  if (agentReference) {
    if (!isSafeRelativePath(agentReference.file)) {
      return err(
        packageError("InvalidAgentReference", "The application agent reference path is invalid.")
      );
    }
    const agentBytes = byName.get(agentReference.file);
    if (agentBytes === undefined) {
      return err(
        packageError(
          "MissingAgentManifest",
          "The application package is missing its declarative agent manifest."
        )
      );
    }
    const agentResult = parseJsonObject(agentBytes, agentReference.file);
    if (agentResult.isErr()) {
      return err(agentResult.error);
    }
    agentManifest = agentResult.value;
  }
  const referenceKind = agentReference ? "declarativeAgentAction" : "agentConnector";
  const bindingSha256 = sha256(sourcePluginBytes);
  const pluginSha256 = sha256(pluginBytes);
  const pluginPath = `${managedPluginPathPrefix}${bindingSha256.slice(0, 16)}.json`;
  const pluginId = `${managedPluginPrefix}${bindingSha256.slice(0, 16)}`;
  const referenceArrays = readReferenceArrays(manifest, agentManifest);
  if (referenceArrays.isErr()) {
    return err(referenceArrays.error);
  }
  const { actions, connectors } = referenceArrays.value;
  const candidates = managedCandidates(actions, connectors);
  if (candidates.length > 1) {
    return err(
      packageError(
        "AmbiguousManagedPluginReference",
        "The application package contains multiple possible managed plugin references."
      )
    );
  }

  const zip = new AdmZip(applicationPackageBytes);
  const zipEntriesToRemove: string[] = [];
  let existingReference = false;
  if (candidates.length === 1) {
    const candidate = candidates[0];
    if (typeof candidate.path !== "string" || !isSafeRelativePath(candidate.path)) {
      return err(
        packageError(
          "AmbiguousManagedPluginReference",
          "The managed plugin reference has an unsafe path."
        )
      );
    }
    const candidateBytes = byName.get(candidate.path);
    if (candidateBytes === undefined) {
      return err(
        packageError(
          candidate.id === pluginId || candidate.path === pluginPath
            ? "ApplicationPackageCollision"
            : "AmbiguousManagedPluginReference",
          "The managed plugin reference does not have a matching package entry."
        )
      );
    }
    if (
      candidate.kind === referenceKind &&
      candidate.id === pluginId &&
      candidate.path === pluginPath
    ) {
      if (
        !candidateBytes.equals(pluginBytes) &&
        !matchesSourcePlugin(candidateBytes, sourcePluginManifest, productId)
      ) {
        return err(
          packageError(
            "ApplicationPackageCollision",
            "The application package contains conflicting managed plugin content."
          )
        );
      }
      existingReference = true;
    } else if (
      isLegacyManagedReference(candidate, candidateBytes) &&
      matchesSourcePlugin(candidateBytes, sourcePluginManifest, productId)
    ) {
      if (candidate.kind === "declarativeAgentAction") {
        if (!agentReference || !agentManifest) {
          return err(
            packageError(
              "AmbiguousManagedPluginReference",
              "The old managed action does not belong to the selected declarative agent."
            )
          );
        }
        agentManifest["actions"] = actions.filter(
          (action) => action["id"] !== candidate.id || action["file"] !== candidate.path
        );
      } else {
        manifest["agentConnectors"] = connectors.filter(
          (connector) => connector["id"] !== candidate.id
        );
      }
      zipEntriesToRemove.push(candidate.path);
    } else {
      return err(
        packageError(
          "ApplicationPackageCollision",
          "The application package contains a conflicting or ambiguous managed plugin reference."
        )
      );
    }
  }

  if (!existingReference && byName.has(pluginPath)) {
    return err(
      packageError(
        "ApplicationPackageCollision",
        "The application package contains a conflicting plugin file."
      )
    );
  }
  for (const entryPath of zipEntriesToRemove) {
    zip.deleteFile(entryPath);
  }

  if (agentReference) {
    if (!agentManifest) {
      return err(
        packageError("MissingAgentManifest", "The application agent manifest is unavailable.")
      );
    }
    if (existingReference) {
      agentManifest["actions"] = actions.map((action) =>
        action["id"] === pluginId && action["file"] === pluginPath
          ? { ...action, id: pluginId, file: pluginPath }
          : action
      );
    } else {
      agentManifest["actions"] = [...actions, { id: pluginId, file: pluginPath }];
    }
    zip.updateFile(
      agentReference.file,
      Buffer.from(JSON.stringify(agentManifest, null, 2), "utf8")
    );
  } else {
    if (!existingReference) {
      manifest["agentConnectors"] = [
        ...connectors,
        {
          id: pluginId,
          displayName: pluginName,
          toolSource: {
            plugin: {
              id: pluginNamespace,
              file: pluginPath,
            },
          },
        },
      ];
    }
    zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest, null, 2), "utf8"));
  }
  if (existingReference) {
    zip.updateFile(pluginPath, pluginBytes);
  } else {
    zip.addFile(pluginPath, pluginBytes);
  }
  const addedPlugin = zip.getEntry(pluginPath);
  if (addedPlugin && !byName.has(pluginPath)) {
    addedPlugin.header.time = new Date("1980-01-01T00:00:00.000Z");
  }
  try {
    const output = zip.toBuffer();
    const outputEntries = readBoundedArchive(output);
    if (outputEntries.isErr()) {
      return err(outputEntries.error);
    }
    return ok({
      bytes: output,
      appId,
      appVersion,
      pluginSha256,
      pluginBindingSha256: bindingSha256,
      managedReference: { kind: referenceKind, id: pluginId, path: pluginPath },
    });
  } catch (error) {
    return err(
      packageError(
        "ApplicationPackagePrepareFailed",
        "The application package could not be prepared.",
        error
      )
    );
  }
}

export function toApplicationPackageError(
  error: SharePointPackageInspectionError
): SharePointApplicationPackageError {
  return packageError(error.code, error.message, error);
}
