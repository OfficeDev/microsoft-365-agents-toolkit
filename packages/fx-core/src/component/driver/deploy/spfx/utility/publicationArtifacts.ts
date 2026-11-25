// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { randomUUID } from "crypto";
import * as fs from "fs-extra";
import path from "path";
import {
  ApiPluginManifestConverter,
  AppManifestUtils,
  DeclarativeAgentManifestConverter,
  TeamsManifestConverter,
} from "@microsoft/teamsfx-api";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";

import {
  assembleApplicationPackage,
  resolvePackageReference,
  SharePointApplicationPackageError,
} from "./applicationPackage";
import {
  inspectSharePointPackage,
  readBoundedArchive,
  SHAREPOINT_PACKAGE_LIMITS,
  SharePointPackageInspectionError,
} from "./packageFile";
import { findUnresolvedTenantValue, resolveTenantOrigin } from "./pluginTenantValues";

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

export interface PreparedPublicationInput {
  projectPath: string;
  packagePath: string;
  appPackagePath: string;
  tenantContext: PreparedPublicationTenantContext;
  signal?: AbortSignal;
}

export interface PreparedPublicationTenantContext {
  tenantId: string;
  tenantUrl: string;
  productId: string;
}

export interface PreparedPublicationOutput {
  sharePointPackagePath: string;
  appPackagePath: string;
}

interface ProjectInput {
  path: string;
  relativePath: string;
  bytes: Buffer;
}

const tenantMcpUrlPlaceholder = /^\{\{\s*TENANT_MCP_URL\s*\}\}$/;
const placeholderPattern = /\{\{\s*([A-Z][A-Z0-9_]*)\s*\}\}/;
const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isWithinProject(root: string, filePath: string): boolean {
  const relative = path.relative(root, filePath);
  return relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function projectRelative(root: string, filePath: string): string {
  return path.relative(root, filePath).split(path.sep).join("/");
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
      return err(
        new PreparedPublicationError("InputTooLarge", "A publication input exceeds safe limits.")
      );
    }

    const chunks: Buffer[] = [];
    let bytesRead = 0;
    while (bytesRead <= maxBytes) {
      const chunk = Buffer.alloc(Math.min(64 * 1024, maxBytes + 1 - bytesRead));
      const result = await fs.read(descriptor, chunk, 0, chunk.length, bytesRead);
      if (result.bytesRead === 0) break;
      chunks.push(chunk.subarray(0, result.bytesRead));
      bytesRead += result.bytesRead;
      if (bytesRead > maxBytes) {
        return err(
          new PreparedPublicationError("InputTooLarge", "A publication input exceeds safe limits.")
        );
      }
    }
    const finalStats = await fs.fstat(descriptor);
    if (finalStats.size !== bytesRead || finalStats.mtimeMs !== stats.mtimeMs) {
      return err(
        new PreparedPublicationError(
          "InputChangedDuringRead",
          "A publication input changed while it was read."
        )
      );
    }
    return ok(Buffer.concat(chunks, bytesRead));
  } catch (error) {
    return err(
      new PreparedPublicationError("InputReadFailed", "A publication input could not be read.", {
        cause: error,
      })
    );
  } finally {
    if (descriptor !== undefined) await fs.close(descriptor);
  }
}

async function readProjectInput(
  root: string,
  inputPath: string
): Promise<Result<ProjectInput, PreparedPublicationError>> {
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
  try {
    const resolved = await fs.realpath(candidate);
    if (!isWithinProject(root, resolved)) {
      return err(
        new PreparedPublicationError(
          "InvalidInputPath",
          "Publication inputs must remain inside the project."
        )
      );
    }
    const read = await readFileBounded(resolved, SHAREPOINT_PACKAGE_LIMITS.inputBytes);
    return read.map((bytes) => ({
      path: resolved,
      relativePath: projectRelative(root, resolved),
      bytes,
    }));
  } catch (error) {
    return err(
      new PreparedPublicationError("InputReadFailed", "A publication input could not be read.", {
        cause: error,
      })
    );
  }
}

function resolvePluginValues(
  pluginBytes: Buffer,
  productId: string,
  tenantContext: PreparedPublicationTenantContext
): Result<Buffer, PreparedPublicationError> {
  let tenantUrl: URL;
  try {
    tenantUrl = new URL(tenantContext.tenantUrl);
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InvalidPreparedTenantContext",
        "The selected tenant URL is invalid.",
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
    productId.toLowerCase() !== tenantContext.productId.toLowerCase()
  ) {
    return err(
      new PreparedPublicationError(
        "InvalidPreparedTenantContext",
        "The selected tenant and SharePoint solution identity do not match."
      )
    );
  }

  let plugin: Record<string, unknown>;
  try {
    const parsed: unknown = JSON.parse(pluginBytes.toString("utf8"));
    if (!isRecord(parsed)) throw new Error("The API plugin root must be an object.");
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
  const endpoint = new URL("/_api/mcp/spfx", tenantUrl.origin);
  endpoint.searchParams.set("solutionId", productId);
  let endpointCount = 0;
  for (const runtime of runtimes) {
    const spec = runtime["spec"];
    if (runtime["type"] !== "RemoteMCPServer" || !isRecord(spec)) continue;
    const url = spec["url"];
    if (typeof url === "string" && tenantMcpUrlPlaceholder.test(url.trim())) {
      spec["url"] = endpoint.toString();
      endpointCount += 1;
    } else if (typeof url === "string" && placeholderPattern.test(url)) {
      return err(
        new PreparedPublicationError(
          "UnresolvedPluginPlaceholder",
          "A registered API endpoint contains an unsupported tenant placeholder."
        )
      );
    }
    if (!resolveTenantOrigin(runtime, tenantUrl.origin)) {
      return err(
        new PreparedPublicationError(
          "UnresolvedPluginPlaceholder",
          "A registered content security policy entry contains an unsupported tenant placeholder."
        )
      );
    }
  }
  if (endpointCount !== 1) {
    return err(
      new PreparedPublicationError(
        "AmbiguousPluginEndpoint",
        "The embedded API plugin must register exactly one tenant-bound SharePoint endpoint."
      )
    );
  }
  const unresolved = findUnresolvedTenantValue(plugin);
  if (unresolved?.kind === "tooDeep") {
    return err(
      new PreparedPublicationError(
        "InvalidPluginManifest",
        "The embedded API plugin is nested too deeply."
      )
    );
  }
  if (unresolved !== undefined) {
    const code =
      unresolved.token === "TENANT_ORIGIN"
        ? "UnqualifiedTenantOriginField"
        : "UnresolvedPluginPlaceholder";
    return err(
      new PreparedPublicationError(
        code,
        unresolved.token === "TENANT_ORIGIN"
          ? "The tenant-origin CSP field is not qualified for this publication format."
          : `The embedded API plugin contains the unresolved placeholder "${unresolved.token.slice(0, 64)}".`
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
    return isRecord(parsed)
      ? ok(parsed)
      : err(
          new PreparedPublicationError(
            "InvalidPackageMetadata",
            `Package metadata "${name}" is invalid.`
          )
        );
  } catch (error) {
    return err(
      new PreparedPublicationError(
        "InvalidPackageMetadata",
        `Package metadata "${name}" is invalid.`,
        { cause: error }
      )
    );
  }
}

function getTeamsSchemaUrl(version: unknown): string | undefined {
  if (typeof version !== "string") return undefined;
  if (["devPreview", "m365DevPreview", "m365devpreview"].includes(version)) {
    return "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json";
  }
  return /^1\.(?:[0-9]|1[0-9]|2[0-9]|30)$/.test(version)
    ? `https://developer.microsoft.com/json-schemas/teams/v${version}/MicrosoftTeams.schema.json`
    : undefined;
}

function getPluginSchemaUrl(version: unknown): string | undefined {
  return typeof version === "string" && /^v2\.[1-4]$/.test(version)
    ? `https://developer.microsoft.com/json-schemas/copilot/plugin/${version}/schema.json`
    : undefined;
}

function getAgentSchemaUrl(version: unknown): string | undefined {
  return typeof version === "string" && /^v1\.[0-9]$/.test(version) && version !== "v1.1"
    ? `https://developer.microsoft.com/json-schemas/copilot/declarative-agent/${version}/schema.json`
    : undefined;
}

async function validateApplicationPackage(
  bytes: Buffer
): Promise<Result<void, PreparedPublicationError>> {
  const entries = readBoundedArchive(bytes);
  if (entries.isErr()) return err(entries.error);
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
  const parsedManifest = jsonObject(manifestBytes, "manifest.json");
  if (parsedManifest.isErr()) return err(parsedManifest.error);
  const manifest = parsedManifest.value;
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
    const teamsSchema = await AppManifestUtils.fetchSchema(teamsSchemaUrl);
    const typedManifest = TeamsManifestConverter.jsonToManifest(manifestBytes.toString("utf8"));
    if ((await AppManifestUtils.validateAgainstSchema(typedManifest, teamsSchema)).length > 0) {
      return err(
        new PreparedPublicationError(
          "ApplicationSchemaValidationFailed",
          "The staged application manifest does not satisfy its supported schema."
        )
      );
    }

    const pluginPaths = new Set<string>();
    const agentPaths: string[] = [];
    for (const containerName of ["copilotAgents", "copilotExtensions"]) {
      const container = manifest[containerName];
      if (!isRecord(container)) continue;
      const references =
        container[containerName === "copilotAgents" ? "declarativeAgents" : "declarativeCopilots"];
      if (!Array.isArray(references)) continue;
      for (const reference of references) {
        if (isRecord(reference) && typeof reference["file"] === "string") {
          const agentPath = resolvePackageReference("manifest.json", reference["file"]);
          if (!agentPath) {
            return err(
              new PreparedPublicationError(
                "ApplicationSchemaValidationFailed",
                "The staged application contains an invalid agent reference."
              )
            );
          }
          agentPaths.push(agentPath);
        }
      }
    }
    const connectors = manifest["agentConnectors"];
    if (Array.isArray(connectors)) {
      for (const connector of connectors) {
        const plugin =
          isRecord(connector) && isRecord(connector["toolSource"])
            ? connector["toolSource"]["plugin"]
            : undefined;
        if (isRecord(plugin) && typeof plugin["file"] === "string") {
          const pluginPath = resolvePackageReference("manifest.json", plugin["file"]);
          if (!pluginPath) {
            return err(
              new PreparedPublicationError(
                "ApplicationSchemaValidationFailed",
                "The staged application contains an invalid plugin reference."
              )
            );
          }
          pluginPaths.add(pluginPath);
        }
      }
    }

    for (const agentPath of agentPaths) {
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
      if (agentObject.isErr()) return err(agentObject.error);
      const schemaUrl = getAgentSchemaUrl(agentObject.value["version"]);
      if (!schemaUrl) {
        return err(
          new PreparedPublicationError(
            "UnsupportedAgentManifestVersion",
            "The staged declarative agent version is not supported."
          )
        );
      }
      const schema = await AppManifestUtils.fetchSchema(schemaUrl);
      const typedAgent = DeclarativeAgentManifestConverter.jsonToManifest(
        agentBytes.toString("utf8")
      );
      if ((await AppManifestUtils.validateAgainstSchema(typedAgent, schema)).length > 0) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "A staged declarative agent does not satisfy its supported schema."
          )
        );
      }
      const actions = agentObject.value["actions"];
      if (Array.isArray(actions)) {
        for (const action of actions) {
          if (isRecord(action) && typeof action["file"] === "string") {
            const pluginPath = resolvePackageReference(agentPath, action["file"]);
            if (!pluginPath) {
              return err(
                new PreparedPublicationError(
                  "ApplicationSchemaValidationFailed",
                  "The staged application contains an invalid API plugin reference."
                )
              );
            }
            pluginPaths.add(pluginPath);
          }
        }
      }
    }

    for (const pluginPath of pluginPaths) {
      const pluginBytes = byName.get(pluginPath);
      const pluginObject = pluginBytes ? jsonObject(pluginBytes, pluginPath) : undefined;
      const schemaUrl =
        pluginObject && pluginObject.isOk()
          ? getPluginSchemaUrl(pluginObject.value["schema_version"])
          : undefined;
      if (!pluginBytes || !pluginObject?.isOk() || !schemaUrl) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "A staged API plugin reference is missing or invalid."
          )
        );
      }
      const schema = await AppManifestUtils.fetchSchema(schemaUrl);
      const typedPlugin = ApiPluginManifestConverter.jsonToManifest(pluginBytes.toString("utf8"));
      if ((await AppManifestUtils.validateAgainstSchema(typedPlugin, schema)).length > 0) {
        return err(
          new PreparedPublicationError(
            "ApplicationSchemaValidationFailed",
            "A staged API plugin does not satisfy its supported schema."
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

async function writeAtomic(filePath: string, bytes: Buffer): Promise<void> {
  const temporaryPath = `${filePath}.${randomUUID()}.tmp`;
  try {
    await fs.writeFile(temporaryPath, bytes, { flag: "wx" });
    await fs.rename(temporaryPath, filePath);
  } catch (error) {
    await fs.remove(temporaryPath);
    throw new PreparedPublicationError(
      "PublicationWriteFailed",
      "A staged publication package could not be written.",
      { cause: error }
    );
  }
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
        "The project directory is unavailable.",
        {
          cause: error,
        }
      )
    );
  }
  const input = await readProjectInput(root, packagePath);
  if (input.isErr()) return err(input.error);
  const inspection = await inspectSharePointPackage(input.value.bytes);
  return inspection.isErr() ? err(inspection.error) : ok(inspection.value.productId);
}

export async function preparePublication(
  input: PreparedPublicationInput
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
        "The project directory is unavailable.",
        {
          cause: error,
        }
      )
    );
  }
  const sourcePackage = await readProjectInput(root, input.packagePath);
  if (sourcePackage.isErr()) return err(sourcePackage.error);
  if (!/^[A-Za-z0-9._-]+\.sppkg$/i.test(path.posix.basename(sourcePackage.value.relativePath))) {
    return err(
      new PreparedPublicationError(
        "InvalidPackageName",
        "The SharePoint package file name is not supported."
      )
    );
  }
  const sourceApplication = await readProjectInput(root, input.appPackagePath);
  if (sourceApplication.isErr()) return err(sourceApplication.error);
  if (input.signal?.aborted) {
    return err(new PreparedPublicationError("PublicationCancelled", "Preparation was cancelled."));
  }

  const sharePointPackage = await inspectSharePointPackage(sourcePackage.value.bytes);
  if (sharePointPackage.isErr()) return err(sharePointPackage.error);
  if (
    sharePointPackage.value.productId.toLowerCase() !== input.tenantContext.productId.toLowerCase()
  ) {
    return err(
      new PreparedPublicationError(
        "SelectedProductMismatch",
        "The selected tenant context does not match the SharePoint solution."
      )
    );
  }
  const plugin = resolvePluginValues(
    sharePointPackage.value.pluginManifest,
    sharePointPackage.value.productId,
    input.tenantContext
  );
  if (plugin.isErr()) return err(plugin.error);

  const application = assembleApplicationPackage(
    sourceApplication.value.bytes,
    plugin.value,
    sharePointPackage.value.pluginManifest,
    sharePointPackage.value.productId
  );
  if (application.isErr()) {
    const error: SharePointApplicationPackageError = application.error;
    return err(new PreparedPublicationError(error.code, error.message, { cause: error }));
  }
  const validation = await validateApplicationPackage(application.value.bytes);
  if (validation.isErr()) return err(validation.error);
  if (input.signal?.aborted) {
    return err(new PreparedPublicationError("PublicationCancelled", "Preparation was cancelled."));
  }

  const stagingRoot = path.resolve(root, ".fx", "sharepoint-publications");
  const stagingDirectory = path.join(stagingRoot, randomUUID());
  try {
    await fs.ensureDir(stagingRoot);
    const resolvedRoot = await fs.realpath(stagingRoot);
    if (!isWithinProject(root, resolvedRoot)) {
      throw new PreparedPublicationError("InvalidStagingPath", "The staging path is invalid.");
    }
    await fs.mkdir(stagingDirectory);
    const resolvedDirectory = await fs.realpath(stagingDirectory);
    if (!isWithinProject(root, resolvedDirectory)) {
      throw new PreparedPublicationError("InvalidStagingPath", "The staging path is invalid.");
    }
    const sharePointPath = path.join(resolvedDirectory, "sharepoint.sppkg");
    const applicationPath = path.join(resolvedDirectory, "application.zip");
    await writeAtomic(sharePointPath, sharePointPackage.value.deploymentBytes);
    if (input.signal?.aborted) {
      throw new PreparedPublicationError("PublicationCancelled", "Preparation was cancelled.");
    }
    await writeAtomic(applicationPath, application.value.bytes);
    if (input.signal?.aborted) {
      throw new PreparedPublicationError("PublicationCancelled", "Preparation was cancelled.");
    }
    return ok({
      sharePointPackagePath: projectRelative(root, sharePointPath),
      appPackagePath: projectRelative(root, applicationPath),
    });
  } catch (error) {
    await fs.remove(stagingDirectory);
    return err(
      error instanceof PreparedPublicationError
        ? error
        : new PreparedPublicationError(
            "PublicationPrepareFailed",
            "The publication packages could not be prepared.",
            { cause: error }
          )
    );
  }
}
