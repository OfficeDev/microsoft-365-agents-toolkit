// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { runOperation, declarativeAgent } from "@microsoft/teamsfx-core-next";
import * as fs from "fs";
import * as path from "path";

export interface AddActionInput {
  projectPath: string;
  agentManifestPath?: string;
  apiSpecPath: string;
  pluginManifestPath: string;
  actionId: string;
}

export async function addActionAction(ctx: AtkContext, input: AddActionInput): Promise<void> {
  const agentManifestPath = await resolveAgentManifestPath(
    input.agentManifestPath,
    input.projectPath
  );

  const result = await runOperation(declarativeAgent.addExistingPluginOp, ctx, {
    agentManifestPath,
    pluginManifestPath: input.pluginManifestPath,
    apiSpecPath: input.apiSpecPath,
    actionId: input.actionId,
  });
  if (result.isErr()) throw toError(result.error);
}

async function resolveAgentManifestPath(
  explicit: string | undefined,
  projectPath: string
): Promise<string> {
  if (explicit) return explicit;
  const result = await declarativeAgent.getAgentManifestPath(projectPath);
  if (result.isErr()) throw toError(result.error);
  return result.value;
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}

// ---------------------------------------------------------------------------
// MCP Action
// ---------------------------------------------------------------------------

export interface AddMCPActionCliInput {
  projectPath: string;
  agentManifestPath?: string;
  pluginManifestPath?: string;
  serverUrl: string;
  serverName: string;
  isLocal: boolean;
  authType: string;
  toolsFilePath?: string;
  selectedTools?: string[];
  oauthAuthUrl?: string;
  oauthTokenUrl?: string;
  oauthRefreshUrl?: string;
}

const validAuthTypes = new Set(["oauth", "api-key", "bearer-token", "microsoft-entra", "none"]);

export async function addMCPActionAction(
  ctx: AtkContext,
  input: AddMCPActionCliInput
): Promise<void> {
  // Validate required fields
  if (!input.serverUrl) {
    throw new Error("--mcp-server-url is required for MCP actions");
  }
  if (!input.serverName) {
    throw new Error("--mcp-server-name is required for MCP actions");
  }
  if (!validAuthTypes.has(input.authType)) {
    throw new Error(
      `Invalid --mcp-auth-type: ${input.authType}. Must be one of: ${[...validAuthTypes].join(", ")}`
    );
  }

  // Load tools from file
  const tools = await loadMCPTools(input.toolsFilePath);

  // Determine selected tools (default to all)
  const selectedTools = input.selectedTools ?? tools.map((t) => t.name);

  // Resolve plugin manifest path — default to <project>/appPackage/ai-plugin.json
  const pluginManifestPath =
    input.pluginManifestPath ?? path.join(input.projectPath, "appPackage", "ai-plugin.json");

  // Build OAuth metadata if auth type is oauth
  const oauthMeta =
    input.authType === "oauth" && input.oauthAuthUrl && input.oauthTokenUrl
      ? {
          authorizationUrl: input.oauthAuthUrl,
          tokenUrl: input.oauthTokenUrl,
          refreshUrl: input.oauthRefreshUrl,
        }
      : undefined;

  const result = await runOperation(declarativeAgent.addMCPActionOp, ctx, {
    pluginManifestPath,
    serverUrlOrIdentifier: input.serverUrl,
    serverName: input.serverName,
    isLocal: input.isLocal,
    auth: input.authType as declarativeAgent.AuthScheme,
    oauthMeta,
    availableTools: tools,
    selectedTools,
  });
  if (result.isErr()) throw toError(result.error);
}

async function loadMCPTools(
  toolsFilePath: string | undefined
): Promise<declarativeAgent.MCPToolDefinition[]> {
  if (!toolsFilePath) {
    return [];
  }

  const resolvedPath = path.resolve(toolsFilePath);
  try {
    const content = await fs.promises.readFile(resolvedPath, "utf-8");
    const parsed: unknown = JSON.parse(content);

    if (!Array.isArray(parsed)) {
      throw new Error("MCP tools file must contain a JSON array of tool definitions");
    }

    return parsed.map((item: Record<string, unknown>, i: number) => {
      if (!item.name || typeof item.name !== "string") {
        throw new Error(`Tool at index ${i} must have a string "name" field`);
      }
      return {
        name: item.name,
        description: typeof item.description === "string" ? item.description : undefined,
        inputSchema:
          item.inputSchema && typeof item.inputSchema === "object"
            ? (item.inputSchema as Record<string, unknown>)
            : undefined,
      };
    });
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") {
      throw new Error(`MCP tools file not found: ${resolvedPath}`);
    }
    throw e;
  }
}
