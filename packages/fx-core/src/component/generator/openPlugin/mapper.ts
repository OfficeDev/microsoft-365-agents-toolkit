// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import { parseAuthor } from "./authorParser";
import { deterministicAppId } from "./deterministicId";
import { toTitleCaseFromKebab, truncateAtWordBoundary } from "./textUtils";
import {
  AuthorizationType,
  ConvertInputs,
  CopyOp,
  MappedManifest,
  OpenPluginMcpServerEntry,
  ParsedOpenPlugin,
} from "./types";

export const MANIFEST_SCHEMA_URL =
  "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json";
export const MANIFEST_VERSION = "devPreview";
export const ACCENT_COLOR = "#4A90D9";

const NAME_SHORT_MAX = 30;
const NAME_FULL_MAX = 100;
const DESC_SHORT_MAX = 80;
const DESC_FULL_MAX = 4000;
const MAX_AGENT_CONNECTORS = 10;

export function mapToTtkProject(parsed: ParsedOpenPlugin, inputs: ConvertInputs): MappedManifest {
  const warnings = [...parsed.warnings];
  const pj = parsed.manifest;
  const pluginName = pj.name;

  if (inputs.packageName !== undefined) {
    warnings.push(
      "--package-name was provided but the devPreview manifest schema does not include 'packageName'; ignored."
    );
  }

  const author = parseAuthor(pj.author);
  const websiteUrl = inputs.websiteUrl ?? pj.homepage ?? author.url;
  if (!websiteUrl) {
    throw new Error(
      "developer.websiteUrl could not be resolved. Set 'homepage' in plugin.json, 'author.url', or pass --website-url."
    );
  }
  if (!inputs.privacyUrl) {
    throw new Error(
      "developer.privacyUrl is required. Pass --privacy-url (Open Plugin spec has no equivalent field)."
    );
  }
  if (!inputs.termsUrl) {
    throw new Error(
      "developer.termsOfUseUrl is required. Pass --terms-url (Open Plugin spec has no equivalent field)."
    );
  }

  const idSeed = inputs.packageName ?? `openplugin:${pluginName}`;
  const appId = inputs.appId ?? deterministicAppId(idSeed);

  const displayName = toTitleCaseFromKebab(pluginName);
  const shortName = truncateAtWordBoundary(displayName, NAME_SHORT_MAX);
  const fullName = truncateAtWordBoundary(displayName, NAME_FULL_MAX);
  const description = pj.description ?? pluginName;
  const shortDesc = truncateAtWordBoundary(description, DESC_SHORT_MAX);
  const fullDesc = truncateAtWordBoundary(description, DESC_FULL_MAX);

  const agentSkills = parsed.skills.map((folder) => ({ folder: `./skills/${folder}` }));

  const agentConnectors = buildAgentConnectors(
    parsed.mcpServers,
    pluginName,
    inputs.defaultAuthType ?? "Auto",
    warnings
  );

  if (agentConnectors.length > MAX_AGENT_CONNECTORS) {
    throw new Error(
      `Too many MCP servers: ${agentConnectors.length}. The manifest caps agentConnectors at ${MAX_AGENT_CONNECTORS}.`
    );
  }

  const developer: Record<string, unknown> = {
    name: author.name ?? "Unknown",
    websiteUrl,
    privacyUrl: inputs.privacyUrl,
    termsOfUseUrl: inputs.termsUrl,
  };

  const manifest: Record<string, unknown> = {
    $schema: MANIFEST_SCHEMA_URL,
    manifestVersion: MANIFEST_VERSION,
    version: pj.version ?? "1.0.0",
    id: appId,
  };
  manifest.developer = developer;
  manifest.name = { short: shortName, full: fullName };
  manifest.description = { short: shortDesc, full: fullDesc };
  manifest.icons = { color: "color.png", outline: "outline.png" };
  manifest.accentColor = ACCENT_COLOR;
  if (agentSkills.length > 0) {
    manifest.agentSkills = agentSkills;
  }
  if (agentConnectors.length > 0) {
    manifest.agentConnectors = agentConnectors;
  }

  const copyOps: CopyOp[] = [];
  if (parsed.skillsRoot) {
    for (const skill of parsed.skills) {
      copyOps.push({
        src: path.join(parsed.skillsRoot, skill),
        destRelative: `appPackage/skills/${skill}`,
      });
    }
  }
  if (parsed.commandsRoot && parsed.commands.length > 0) {
    copyOps.push({
      src: parsed.commandsRoot,
      destRelative: "appPackage/commands",
    });
  }

  return { manifest, copyOps, warnings };
}

function buildAgentConnectors(
  mcpServers: Record<string, OpenPluginMcpServerEntry>,
  pluginName: string,
  defaultAuth: "Auto" | AuthorizationType,
  warnings: string[]
): Record<string, unknown>[] {
  const out: Record<string, unknown>[] = [];
  const serverNames = Object.keys(mcpServers).sort();
  for (const name of serverNames) {
    const server = mcpServers[name];
    const url = typeof server.url === "string" ? server.url.trim() : "";
    if (!url) {
      warnings.push(
        `Skipping MCP server '${name}': no URL found (stdio servers require manual localMcpServer configuration).`
      );
      continue;
    }
    const authType = resolveAuthType(url, defaultAuth);
    const authorization: Record<string, unknown> = { type: authType };
    if (authType !== "None") {
      authorization.referenceId = `${pluginName}-${name}-auth`;
    }
    const description =
      typeof server.description === "string" && server.description
        ? server.description
        : `Remote MCP server providing tools for ${pluginName}`;
    out.push({
      id: name,
      displayName: `${name} MCP Server`,
      description,
      toolSource: {
        remoteMcpServer: {
          mcpServerUrl: url,
          authorization,
        },
      },
    });
  }
  return out;
}

function resolveAuthType(url: string, defaultAuth: "Auto" | AuthorizationType): AuthorizationType {
  if (defaultAuth !== "Auto") {
    return defaultAuth;
  }
  const isHttps = /^https:\/\//i.test(url);
  const isLocal = /localhost|127\.0\.0\.1/.test(url);
  if (isHttps && !isLocal) {
    return "OAuthPluginVault";
  }
  return "None";
}
