// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import * as dotenv from "dotenv";
import { parse as parseYaml } from "yaml";

type RegistrationKind = "oauth-registration" | "dcr-registration";
type ResourceKind =
  | RegistrationKind
  | "teams-app"
  | "m365-title"
  | "aad-app"
  | "service-principal"
  | "monitoring-app";

interface ResourceOwnership {
  kind: ResourceKind;
  envKey: string;
  existedBefore: boolean;
}

interface McpAuthOwnership {
  projectName: string;
  runId: string;
  mode: "static-oauth" | "dcr" | "entra-sso";
  sharedRegistrationId?: string;
  resources: ResourceOwnership[];
}

export interface McpTeardownOperations {
  deleteTeamsApp(id: string): Promise<void>;
  unacquireM365Title(id: string): Promise<void>;
  deleteRegistration?(kind: RegistrationKind, id: string): Promise<void>;
}

export interface McpTeardownResult {
  deleted: { kind: ResourceKind; id: string }[];
  skipped: { kind: ResourceKind; reason: string }[];
  retained: { kind: ResourceKind; id: string; reason: string }[];
}

function isMcpAuthRun(): boolean {
  return (
    !!process.env.MCP_AUTH_RUN_ID ||
    ["static-oauth", "dcr", "entra-sso"].includes(
      process.env.MCP_AUTH_MODE ?? "",
    )
  );
}

export function skipMcpSharedIdentityCleanup(): boolean {
  if (!isMcpAuthRun()) return false;
  console.info(
    "MCP teardown: skipping mini-managed AAD/SP/monitoring identity cleanup.",
  );
  return true;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isResourceKind(value: unknown): value is ResourceKind {
  return (
    typeof value === "string" &&
    [
      "teams-app",
      "m365-title",
      "oauth-registration",
      "dcr-registration",
      "aad-app",
      "service-principal",
      "monitoring-app",
    ].includes(value)
  );
}

async function isMcpAuthProject(projectPath: string): Promise<boolean> {
  let configSource: string;
  try {
    configSource = await fs.readFile(
      path.join(projectPath, ".vscode", "mcp.json"),
      "utf8",
    );
  } catch (error: unknown) {
    if (isRecord(error) && error.code === "ENOENT") return false;
    throw error;
  }
  const config: unknown = parseYaml(configSource);
  if (!isRecord(config) || !isRecord(config.servers)) return false;
  const resourceUrls = Object.values(config.servers).flatMap((server) =>
    isRecord(server) &&
    typeof server.url === "string" &&
    /^https?:\/\//.test(server.url)
      ? [server.url]
      : [],
  );
  if (!resourceUrls.length) return false;
  const workflow: unknown = parseYaml(
    await fs.readFile(path.join(projectPath, "m365agents.yml"), "utf8"),
  );
  return (
    isRecord(workflow) &&
    Array.isArray(workflow.provision) &&
    workflow.provision.some((action: unknown) => {
      if (!isRecord(action) || !isRecord(action.with)) return false;
      const resourceUrl =
        action.uses === "dcr/register"
          ? action.with.mcpResourceUrl
          : action.uses === "oauth/register"
            ? action.with.baseUrl
            : undefined;
      return (
        typeof resourceUrl === "string" && resourceUrls.includes(resourceUrl)
      );
    })
  );
}

function parseOwnership(value: unknown): McpAuthOwnership {
  if (
    !isRecord(value) ||
    value.version !== 1 ||
    typeof value.projectName !== "string" ||
    !value.projectName ||
    typeof value.runId !== "string" ||
    !value.runId ||
    (value.mode !== "static-oauth" &&
      value.mode !== "dcr" &&
      value.mode !== "entra-sso") ||
    !Array.isArray(value.resources) ||
    (value.sharedRegistrationId !== undefined &&
      (value.mode !== "entra-sso" ||
        typeof value.sharedRegistrationId !== "string"))
  ) {
    throw new Error("Invalid MCP authentication resource ownership record.");
  }
  const resources: ResourceOwnership[] = [];
  for (const resource of value.resources) {
    if (
      !isRecord(resource) ||
      !isResourceKind(resource.kind) ||
      typeof resource.envKey !== "string" ||
      !/^[A-Z][A-Z0-9_]*$/.test(resource.envKey) ||
      typeof resource.existedBefore !== "boolean" ||
      (resource.kind === "dcr-registration" && value.mode !== "dcr") ||
      (resource.kind === "oauth-registration" && value.mode === "dcr")
    ) {
      throw new Error(
        "Invalid resource in MCP authentication ownership record.",
      );
    }
    resources.push({
      kind: resource.kind,
      envKey: resource.envKey,
      existedBefore: resource.existedBefore,
    });
  }
  return {
    projectName: value.projectName,
    runId: value.runId,
    mode: value.mode,
    sharedRegistrationId: value.sharedRegistrationId,
    resources,
  };
}

export async function tryMcpAuthTeardown(
  projectPath: string,
  operations: McpTeardownOperations,
  runId = process.env.MCP_AUTH_RUN_ID,
): Promise<McpTeardownResult | undefined> {
  let source: string;
  try {
    source = await fs.readFile(
      path.join(projectPath, ".auth-test", "ownership.json"),
      "utf8",
    );
  } catch (error: unknown) {
    if (isRecord(error) && error.code === "ENOENT") {
      if (runId || isMcpAuthRun() || (await isMcpAuthProject(projectPath))) {
        throw new Error(
          "Missing ownership record for the MCP authentication run.",
        );
      }
      return undefined;
    }
    throw error;
  }
  const ownership = parseOwnership(JSON.parse(source));
  if (ownership.projectName !== path.basename(path.resolve(projectPath))) {
    throw new Error("MCP ownership record belongs to a different project.");
  }
  const result: McpTeardownResult = { deleted: [], skipped: [], retained: [] };
  const environment = dotenv.parse(
    await fs.readFile(path.join(projectPath, "env", ".env.dev")),
  );
  const expectedRunId = runId ?? environment.MCP_AUTH_RUN_ID;
  if (!expectedRunId || ownership.runId !== expectedRunId) {
    result.skipped = ownership.resources.map(({ kind }) => ({
      kind,
      reason: "foreign-run",
    }));
    return result;
  }
  const ownsTeamsApp = ownership.resources.some(
    (resource) =>
      resource.kind === "teams-app" &&
      !resource.existedBefore &&
      !!environment[resource.envKey],
  );
  for (const resource of ownership.resources) {
    const id = environment[resource.envKey];
    let reason: string | undefined;
    if (
      resource.kind === "aad-app" ||
      resource.kind === "service-principal" ||
      resource.kind === "monitoring-app"
    ) {
      reason = "mini-managed-identity";
    } else if (!id) {
      reason = "not-created";
    } else if (
      resource.existedBefore ||
      id === ownership.sharedRegistrationId
    ) {
      reason = "pre-existing-resource";
    } else if (resource.kind === "m365-title" && !ownsTeamsApp) {
      reason = "unproven-parent-app-ownership";
    }
    if (reason) {
      result.skipped.push({ kind: resource.kind, reason });
      continue;
    }
    if (resource.kind === "teams-app") {
      await operations.deleteTeamsApp(id);
    } else if (resource.kind === "m365-title") {
      await operations.unacquireM365Title(id);
    } else if (
      resource.kind === "oauth-registration" ||
      resource.kind === "dcr-registration"
    ) {
      if (!operations.deleteRegistration) {
        result.retained.push({
          kind: resource.kind,
          id,
          reason: "unsupported-cleanup-api",
        });
        continue;
      }
      await operations.deleteRegistration(resource.kind, id);
    } else {
      throw new Error("MCP teardown cannot delegate shared identity deletion.");
    }
    result.deleted.push({ kind: resource.kind, id });
  }
  return result;
}
