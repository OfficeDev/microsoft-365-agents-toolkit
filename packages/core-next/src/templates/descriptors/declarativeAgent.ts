// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";
import { mcpServerUrlQuestion, officeAddinFolderQuestion } from "../../questions/commonQuestions";
import { copyExistMetaOSProject, extendToDA, unifyProjectID } from "../helpers/metaOSHelper";
import { ok, err } from "neverthrow";

/**
 * Declarative Agent template names matching the template repository folder names.
 * These must match the directory names inside template zip archives
 * (e.g. `templates/vsc/common/declarative-agent-basic/`).
 */
export const DATemplateNames = {
  Basic: "declarative-agent-basic",
  ActionFromScratch: "declarative-agent-with-action-from-scratch",
  ActionFromScratchOAuth: "declarative-agent-with-action-from-scratch-oauth",
  ActionFromScratchBearer: "declarative-agent-with-action-from-scratch-bearer",
  /** Entra SSO reuses the same folder as ActionFromScratch with MicrosoftEntra flag */
  ActionFromScratchEntra: "declarative-agent-with-action-from-scratch",
  ExistingAction: "declarative-agent-with-action-from-existing-api",
  TypeSpec: "declarative-agent-typespec",
  MCP: "declarative-agent-with-action-from-mcp",
  /** MCP Local reuses the MCP folder with IsLocalMCP flag */
  MCPLocal: "declarative-agent-with-action-from-mcp",
  GraphConnector: "graph-connector",
  MetaOS: "declarative-agent-meta-os-new-project",
  MetaOSUpgrade: "declarative-agent-meta-os-upgrade-project",
} as const;

/**
 * Create a standard DA scaffold function for a given template name.
 * This is a factory — all DA templates share the same scaffold pipeline,
 * varying only in template name, language, and replace map.
 */
function makeDAScaffoldFn(templateName: string, scaffoldLang?: string) {
  return async (ctx: AtkContext, opts: TemplateActionOptions) => {
    const replaceMap: Record<string, string> = {
      ...getTemplateReplaceMap({
        appName: opts.projectName,
        ...opts,
      }),
      DeclarativeCopilot: "true",
    };

    // Add auth-specific variables
    if (opts.authType === "microsoft-entra") {
      replaceMap.MicrosoftEntra = "true";
    }
    if (opts.isLocalMCP) {
      replaceMap.IsLocalMCP = "true";
    }
    if (typeof opts.mcpServerUrl === "string") {
      replaceMap.MCPForDAServerUrl = opts.mcpServerUrl;
    }

    const tplInfo: TemplateInfo = {
      templateName,
      language: scaffoldLang ?? convertToLangKey(opts.language),
      replaceMap,
      filterFn: opts.filterFn as ((fileName: string) => boolean) | undefined,
    };

    const result = await scaffoldTemplates(ctx, [tplInfo], opts.destinationPath);
    return result.map((files) => ({
      projectPath: opts.destinationPath,
      warnings: files.length === 0 ? ["No files were scaffolded"] : undefined,
    }));
  };
}

/**
 * All Declarative Agent template descriptors.
 */
export const daTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "da/basic",
    name: "Declarative Agent (Basic)",
    description: "A basic declarative agent with no plugins",
    category: "declarative-agent",
    languages: ["common", "csharp"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.Basic),
    displayOrder: 1,
    tags: ["teamsApp", "publishable", "declarativeAgent"],
  },
  {
    id: "da/api-plugin-no-auth",
    name: "Declarative Agent with API Plugin (No Auth)",
    description: "API plugin with no authentication",
    category: "declarative-agent",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.ActionFromScratch),
    displayOrder: 2,
    tags: ["teamsApp", "publishable", "declarativeAgent", "apiPlugin"],
  },
  {
    id: "da/api-plugin-oauth",
    name: "Declarative Agent with API Plugin (OAuth)",
    description: "API plugin with OAuth authentication",
    category: "declarative-agent",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.ActionFromScratchOAuth),
    displayOrder: 3,
    tags: ["teamsApp", "publishable", "declarativeAgent", "apiPlugin", "apiAuthOAuth", "aad"],
  },
  {
    id: "da/api-plugin-bearer",
    name: "Declarative Agent with API Plugin (Bearer)",
    description: "API plugin with bearer token authentication",
    category: "declarative-agent",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.ActionFromScratchBearer),
    displayOrder: 4,
    tags: ["teamsApp", "publishable", "declarativeAgent", "apiPlugin", "apiAuthApiKey"],
  },
  {
    id: "da/api-plugin-entra-sso",
    name: "Declarative Agent with API Plugin (Entra SSO)",
    description: "API plugin with Entra SSO authentication",
    category: "declarative-agent",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.ActionFromScratchEntra),
    displayOrder: 5,
    tags: ["teamsApp", "publishable", "declarativeAgent", "apiPlugin", "apiAuthOAuth"],
  },
  {
    id: "da/existing-action",
    name: "Declarative Agent with Existing Action",
    description: "Use an existing API plugin action",
    category: "declarative-agent",
    languages: ["common"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.ExistingAction),
    displayOrder: 6,
    tags: ["teamsApp", "publishable", "declarativeAgent", "apiPlugin"],
  },
  {
    id: "da/typespec",
    name: "Declarative Agent from TypeSpec",
    description: "Generate API plugin from a TypeSpec definition",
    category: "declarative-agent",
    languages: ["typescript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.TypeSpec, "common"),
    displayOrder: 7,
    // typeSpec/compile driver not yet ported to core-next — skip E2E tests
    testable: false,
  },
  {
    id: "da/mcp-remote",
    name: "Declarative Agent with MCP Server (Remote)",
    description: "Connect to a remote MCP server",
    category: "declarative-agent",
    languages: ["common"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.MCP),
    displayOrder: 8,
    questions: [mcpServerUrlQuestion()],
    // Template ships with empty ai-plugin.json (functions/runtimes populated
    // when MCP server is configured) — MOS sideloading rejects empty plugins
    testable: false,
  },
  {
    id: "da/mcp-local",
    name: "Declarative Agent with MCP Server (Local)",
    description: "Run a local MCP server alongside the agent",
    category: "declarative-agent",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.MCPLocal, "common"),
    displayOrder: 9,
    // Requires odr.exe + empty ai-plugin.json (same MOS rejection as mcp-remote)
    testable: false,
  },
  {
    id: "da/graph-connector",
    name: "Declarative Agent with Graph Connector",
    description: "Include a Microsoft Graph connector for custom data",
    category: "declarative-agent",
    languages: ["typescript"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.GraphConnector),
    displayOrder: 10,
    testable: false,
  },
  {
    id: "da/metaos",
    name: "Declarative Agent (MetaOS/WXP)",
    description: "Cross-platform agent for Word, Excel, and PowerPoint",
    category: "declarative-agent",
    languages: ["common"],
    scaffoldFn: makeDAScaffoldFn(DATemplateNames.MetaOS),
    displayOrder: 11,
    featureFlag: "DAMetaOS",
    tags: ["teamsApp", "publishable", "declarativeAgent"],
  },
  {
    id: "da/metaos-upgrade",
    name: "Upgrade MetaOS Add-in to DA",
    description: "Upgrade an existing MetaOS Office Add-in project to a Declarative Agent",
    category: "declarative-agent",
    languages: ["common"],
    scaffoldFn: async (_ctx: AtkContext, opts: TemplateActionOptions) => {
      const sourceFolder = opts.officeAddinFolder as string | undefined;
      if (!sourceFolder) {
        return ok({
          projectPath: opts.destinationPath,
          warnings: ["No source project folder specified"],
        });
      }
      await copyExistMetaOSProject(sourceFolder, opts.destinationPath);
      await extendToDA(opts.destinationPath, opts.projectName);

      // Overlay the upgrade template files (README, m365agents.yml, env)
      const replaceMap = getTemplateReplaceMap({ appName: opts.projectName, ...opts });
      const tplInfo: TemplateInfo = {
        templateName: DATemplateNames.MetaOSUpgrade,
        language: "common",
        replaceMap,
      };
      const result = await scaffoldTemplates(_ctx, [tplInfo], opts.destinationPath);
      if (result.isErr()) return err(result.error);

      await unifyProjectID(opts.destinationPath);
      return ok({ projectPath: opts.destinationPath });
    },
    displayOrder: 12,
    featureFlag: "DAMetaOS",
    // Lifecycle E2E blocked by DAMetaOS feature flag.
    // Enable once the flag is on in CI.
    testable: false,
    questions: [officeAddinFolderQuestion()],
  },
];
