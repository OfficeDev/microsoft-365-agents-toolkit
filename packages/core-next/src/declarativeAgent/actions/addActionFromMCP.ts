// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { Result, ok, err } from "neverthrow";
import { APIPluginManifestWrapper, RuntimeType } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { userError, systemError } from "../../core/error";
import type { AddMCPActionInput, MCPToolDefinition } from "../types";

/** Prefix for local MCP server endpoints. */
const LOCAL_MCP_PREFIX = "localmcp://";

/**
 * Add or update an MCP-backed action in the plugin manifest (ai-plugin.json).
 *
 * Handles both local (stdio) and remote (HTTP) MCP servers.
 * For remote servers with pre-fetched tools, writes a mcp-tools.json sidecar file.
 */
export async function addMCPAction(input: AddMCPActionInput): Promise<Result<void, AtkError>> {
  const {
    pluginManifestPath,
    serverUrlOrIdentifier,
    serverName,
    isLocal,
    auth,
    oauthMeta: _oauthMeta,
    availableTools,
    selectedTools,
  } = input;

  try {
    if (!(await fileExists(pluginManifestPath))) {
      return err(
        userError("PluginManifestNotFound", `Plugin manifest not found: ${pluginManifestPath}`, {
          source: "declarativeAgent/actions",
        })
      );
    }

    const wrapper = await APIPluginManifestWrapper.read(pluginManifestPath);

    // Filter tools to only the selected ones
    const selectedToolDefs = availableTools.filter((t) => selectedTools.includes(t.name));

    // Remove previous tools for the same server URL to avoid duplicates
    const serverUrl = isLocal
      ? `${LOCAL_MCP_PREFIX}${serverUrlOrIdentifier}`
      : serverUrlOrIdentifier;
    removePreviousToolsForServer(wrapper, serverUrl);

    // Add selected tool functions
    for (const tool of selectedToolDefs) {
      wrapper.addFunction(tool.name, tool.description ?? "");
    }

    // Build runtime entry
    const runForFunctions = selectedToolDefs.map((t) => t.name);
    const endpoint = isLocal
      ? `${LOCAL_MCP_PREFIX}${serverUrlOrIdentifier}`
      : serverUrlOrIdentifier;

    // For remote MCP with pre-fetched tools, write mcp-tools.json sidecar
    if (!isLocal && selectedToolDefs.length > 0) {
      const toolsFilePath = path.join(path.dirname(pluginManifestPath), "mcp-tools.json");
      const toolsContent = buildToolsDescription(selectedToolDefs);
      await fs.promises.writeFile(toolsFilePath, JSON.stringify(toolsContent, null, 2), "utf-8");
    }

    // Build and add runtime via addRuntime for full control over type/spec/auth
    const runtimeType = isLocal ? RuntimeType.LocalPlugin : RuntimeType.RemoteMCPServer;
    const runtime: Record<string, unknown> = {
      type: runtimeType,
      spec: { url: endpoint },
      run_for_functions: runForFunctions,
    };
    if (!isLocal && selectedToolDefs.length > 0) {
      (runtime.spec as Record<string, unknown>).mcp_tool_description = { file: "mcp-tools.json" };
    }
    if (auth === "oauth" || auth === "microsoft-entra") {
      runtime.auth = { type: "OAuthPluginVault", reference_id: `${serverName}_oauth` };
    } else {
      runtime.auth = { type: "None" };
    }
    wrapper.addRuntime(runtime as never);

    await wrapper.save(pluginManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("AddMCPActionFailed", `Failed to add MCP action: ${e}`, {
        source: "declarativeAgent/actions",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Remove tools (functions + runtime entries) associated with a server URL.
 */
function removePreviousToolsForServer(wrapper: APIPluginManifestWrapper, serverUrl: string): void {
  // Find runtimes matching this server URL and collect their run_for_functions
  const functionsToRemove = new Set<string>();
  for (const rt of wrapper.runtimes) {
    const spec = rt as { spec?: { url?: string }; run_for_functions?: string[] };
    if (spec.spec?.url === serverUrl) {
      for (const fn of spec.run_for_functions ?? []) {
        functionsToRemove.add(fn);
      }
    }
  }

  // Remove the functions and the runtime
  for (const fn of functionsToRemove) {
    wrapper.removeFunction(fn);
  }
  wrapper.removeRuntimeBySpecUrl(serverUrl);
}

/**
 * Build the mcp-tools.json content for pre-fetched tools.
 */
function buildToolsDescription(
  tools: MCPToolDefinition[]
): Array<{ name: string; description?: string; inputSchema?: Record<string, unknown> }> {
  return tools.map((t) => ({
    name: t.name,
    ...(t.description && { description: t.description }),
    ...(t.inputSchema && { inputSchema: t.inputSchema }),
  }));
}

async function fileExists(p: string): Promise<boolean> {
  try {
    await fs.promises.access(p);
    return true;
  } catch {
    return false;
  }
}
