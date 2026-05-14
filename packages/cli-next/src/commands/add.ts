// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandlerWithContext } from "../handler";
import { addActionAction, addMCPActionAction } from "../actions/addAction";
import { addCapabilityAction } from "../actions/addCapability";
import { addAuthConfigAction } from "../actions/addAuthConfig";

/**
 * Add feature commands: add action, add capability, add auth-config
 */
export function createAddCommands(program: Command): void {
  const add = program.command("add").description("Add a feature or capability to your project");

  // atk add action (was "add plugin" in old CLI)
  // Supports two modes via --api-plugin-type: api-spec (default) and mcp
  add
    .command("action")
    .description("Add an action to extend a Declarative Agent")
    .option(
      "--api-plugin-type <type>",
      "Action type: api-spec (OpenAPI) or mcp (MCP server)",
      "api-spec"
    )
    // OpenAPI options (required when api-plugin-type = api-spec)
    .option("--api-spec-path <path>", "Path to the OpenAPI specification file")
    .option("--plugin-manifest-path <path>", "Path to the plugin manifest file")
    .option("--action-id <id>", "Unique identifier for the action")
    // MCP options (required when api-plugin-type = mcp)
    .option("--mcp-server-url <url>", "MCP server URL or local identifier")
    .option("--mcp-server-name <name>", "Human-readable MCP server name")
    .option("--mcp-is-local", "MCP server is local (stdio) vs remote (HTTP)", false)
    .option(
      "--mcp-auth-type <type>",
      "MCP auth type: oauth, api-key, bearer-token, microsoft-entra, none",
      "none"
    )
    .option("--mcp-tools-file <path>", "Path to JSON file with MCP tool definitions")
    .option("--mcp-selected-tools <names...>", "Subset of tool names to include")
    .option("--mcp-oauth-auth-url <url>", "OAuth authorization URL (for mcp auth=oauth)")
    .option("--mcp-oauth-token-url <url>", "OAuth token URL (for mcp auth=oauth)")
    .option("--mcp-oauth-refresh-url <url>", "OAuth refresh URL (for mcp auth=oauth)")
    // Shared options
    .option("--agent-manifest-path <path>", "Path to the DA manifest (auto-detected if omitted)")
    .action(
      wrapHandlerWithContext("add action", async (ctx, opts) => {
        const pluginType = opts.apiPluginType as string;
        const projectPath = ctx.projectPath ?? process.cwd();

        if (pluginType === "mcp") {
          await addMCPActionAction(ctx, {
            projectPath,
            agentManifestPath: opts.agentManifestPath as string | undefined,
            pluginManifestPath: opts.pluginManifestPath as string | undefined,
            serverUrl: opts.mcpServerUrl as string,
            serverName: opts.mcpServerName as string,
            isLocal: opts.mcpIsLocal as boolean,
            authType: (opts.mcpAuthType as string) ?? "none",
            toolsFilePath: opts.mcpToolsFile as string | undefined,
            selectedTools: opts.mcpSelectedTools as string[] | undefined,
            oauthAuthUrl: opts.mcpOauthAuthUrl as string | undefined,
            oauthTokenUrl: opts.mcpOauthTokenUrl as string | undefined,
            oauthRefreshUrl: opts.mcpOauthRefreshUrl as string | undefined,
          });
          console.log("MCP action added successfully.");
        } else {
          // Validate required options for api-spec mode
          if (!opts.apiSpecPath || !opts.pluginManifestPath || !opts.actionId) {
            throw new Error(
              "Options --api-spec-path, --plugin-manifest-path, and --action-id are required for api-spec actions"
            );
          }
          await addActionAction(ctx, {
            projectPath,
            agentManifestPath: opts.agentManifestPath as string | undefined,
            apiSpecPath: opts.apiSpecPath as string,
            pluginManifestPath: opts.pluginManifestPath as string,
            actionId: opts.actionId as string,
          });
          console.log("Action added successfully.");
        }
      })
    );

  // atk add capability
  add
    .command("capability")
    .description("Add a capability to extend Copilot (knowledge, web-search, etc.)")
    .requiredOption(
      "--capability-type <type>",
      "Capability type: web-search, onedrive-sharepoint, graph-connector, embedded-knowledge"
    )
    .option("--agent-manifest-path <path>", "Path to the DA manifest (auto-detected if omitted)")
    .option("--site-url <url>", "SharePoint site URL (for onedrive-sharepoint)")
    .option("--connection-ids <ids...>", "Graph Connector connection IDs (for graph-connector)")
    .option("--file-paths <paths...>", "File paths for embedded knowledge")
    .action(
      wrapHandlerWithContext("add capability", async (ctx, opts) => {
        await addCapabilityAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          agentManifestPath: opts.agentManifestPath as string | undefined,
          source: opts.capabilityType as
            | "web-search"
            | "onedrive-sharepoint"
            | "graph-connector"
            | "embedded-knowledge",
          siteUrl: opts.siteUrl as string | undefined,
          connectionIds: opts.connectionIds as string[] | undefined,
          filePaths: opts.filePaths as string[] | undefined,
        });
        console.log("Capability added successfully.");
      })
    );

  // atk add auth-config
  add
    .command("auth-config")
    .description("Add authentication configuration to a Declarative Agent")
    .requiredOption("--auth-type <type>", "Authentication type: oauth, api-key")
    .requiredOption("--auth-name <name>", "Name for the auth configuration")
    .requiredOption("--spec-path <path>", "Relative path to the API spec file")
    .option("--yml-path <path>", "Path to teamsapp.yml (defaults to <project>/teamsapp.yml)")
    .option("--entra", "Use Microsoft Entra for OAuth", false)
    .option("--enable-pkce", "Enable PKCE for OAuth", false)
    .option("--registration-id <id>", "Existing registration ID to reuse")
    .action(
      wrapHandlerWithContext("add auth-config", async (ctx, opts) => {
        await addAuthConfigAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          authType: opts.authType as "oauth" | "api-key",
          ymlPath: opts.ymlPath as string | undefined,
          authName: opts.authName as string,
          specPath: opts.specPath as string,
          entra: opts.entra as boolean | undefined,
          enablePkce: opts.enablePkce as boolean | undefined,
          registrationId: opts.registrationId as string | undefined,
        });
        console.log("Auth configuration added successfully.");
      })
    );
}
