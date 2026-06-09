// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import * as path from "path";
import {
  AtkExtensionBlock,
  OpenPluginManifest,
  OpenPluginMcpJson,
  OpenPluginMcpServerEntry,
  ParsedOpenPlugin,
} from "./types";

interface ManifestLocation {
  relPath: string;
  kind: "open-plugin" | "claude-plugin" | "cursor-plugin";
}

const MANIFEST_LOCATIONS: ManifestLocation[] = [
  { relPath: ".plugin/plugin.json", kind: "open-plugin" },
  { relPath: ".claude-plugin/plugin.json", kind: "claude-plugin" },
  { relPath: ".cursor-plugin/plugin.json", kind: "cursor-plugin" },
];

const SKILL_NAME_RE = /^[a-z0-9][a-z0-9-_]*$/;

const UNMAPPED_FIELDS: Array<keyof OpenPluginManifest> = [
  "agents",
  "hooks",
  "rules",
  "lspServers",
  "outputStyles",
];

function requireStringPathOverride(value: unknown, field: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value === "string") {
    return value;
  }
  throw new Error(
    `Open Plugin '${field}' override is set to a non-string value. Only the single-string form ` +
      `(e.g. \"${field}\": \"./custom/path\") is supported by this converter today.`
  );
}

export async function readOpenPluginDir(root: string): Promise<ParsedOpenPlugin> {
  const warnings: string[] = [];
  const absRoot = path.resolve(root);
  if (!(await fs.pathExists(absRoot))) {
    throw new Error(`Plugin directory not found: ${absRoot}`);
  }

  // 1. Probe manifest locations.
  let manifestPath: string | undefined;
  let manifestKind: ManifestLocation["kind"] | undefined;
  for (const loc of MANIFEST_LOCATIONS) {
    const candidate = path.join(absRoot, loc.relPath);
    if (await fs.pathExists(candidate)) {
      manifestPath = candidate;
      manifestKind = loc.kind;
      break;
    }
  }
  if (!manifestPath || !manifestKind) {
    throw new Error(
      `No Open Plugin manifest found in ${absRoot}. Looked for: ` +
        MANIFEST_LOCATIONS.map((l) => l.relPath).join(", ")
    );
  }
  const manifest = (await fs.readJSON(manifestPath)) as OpenPluginManifest;
  if (!manifest.name || typeof manifest.name !== "string") {
    throw new Error(`plugin.json is missing required 'name' field at ${manifestPath}`);
  }

  // 2. MCP servers (with optional path override).
  const mcpRel = requireStringPathOverride(manifest.mcpServers, "mcpServers") ?? ".mcp.json";
  const mcpAbs = path.resolve(absRoot, mcpRel);
  const mcpServers: Record<string, OpenPluginMcpServerEntry> = {};
  if (await fs.pathExists(mcpAbs)) {
    const mcpJson = (await fs.readJSON(mcpAbs)) as OpenPluginMcpJson;
    const source =
      mcpJson &&
      typeof mcpJson === "object" &&
      mcpJson.mcpServers &&
      typeof mcpJson.mcpServers === "object"
        ? mcpJson.mcpServers
        : mcpJson;
    if (source && typeof source === "object") {
      for (const [name, value] of Object.entries(source)) {
        if (value && typeof value === "object" && !Array.isArray(value)) {
          mcpServers[name] = value as OpenPluginMcpServerEntry;
        }
      }
    }
  }

  // 3. Skills.
  const skillsRel = requireStringPathOverride(manifest.skills, "skills") ?? "skills";
  const skillsAbs = path.resolve(absRoot, skillsRel);
  const skills: string[] = [];
  let skillsRoot: string | undefined;
  if (await fs.pathExists(skillsAbs)) {
    skillsRoot = skillsAbs;
    const entries = await fs.readdir(skillsAbs, { withFileTypes: true });
    for (const entry of entries) {
      if (!entry.isDirectory()) {
        continue;
      }
      if (!SKILL_NAME_RE.test(entry.name)) {
        warnings.push(
          `Skipping skill folder '${entry.name}': name does not match ${SKILL_NAME_RE.source}.`
        );
        continue;
      }
      const skillMd = path.join(skillsAbs, entry.name, "SKILL.md");
      if (await fs.pathExists(skillMd)) {
        skills.push(entry.name);
      }
    }
    skills.sort();
  }

  // 4. Commands.
  const commandsRel = requireStringPathOverride(manifest.commands, "commands") ?? "commands";
  const commandsAbs = path.resolve(absRoot, commandsRel);
  const commands: string[] = [];
  let commandsRoot: string | undefined;
  if (await fs.pathExists(commandsAbs)) {
    commandsRoot = commandsAbs;
    const entries = await fs.readdir(commandsAbs, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.isFile() && entry.name.toLowerCase().endsWith(".md")) {
        commands.push(entry.name);
      }
    }
    commands.sort();
  }

  // 5. Unmapped Open Plugin fields → emit a warning so the caller knows we
  // dropped them. We do not throw — the spec says hosts may ignore unknown
  // component types.
  const manifestRecord = manifest as unknown as Record<string, unknown>;
  for (const field of UNMAPPED_FIELDS) {
    if (manifestRecord[field] !== undefined) {
      warnings.push(
        `Open Plugin '${field}' field is present but not supported by MOS3 today; dropped.`
      );
    }
  }

  // 6. Icons.
  const hasColorPng = await fs.pathExists(path.join(absRoot, "color.png"));
  const hasOutlinePng = await fs.pathExists(path.join(absRoot, "outline.png"));

  // 7. Round-trip extension block (written by `atk export openplugin`).
  const atkExtension = readAtkExtensionBlock(manifest);

  return {
    pluginRoot: absRoot,
    manifest,
    manifestPath,
    manifestKind,
    mcpServers,
    skills,
    skillsRoot,
    commands,
    commandsRoot,
    hasColorPng,
    hasOutlinePng,
    warnings,
    atkExtension,
  };
}

const ATK_EXTENSION_KEY = "x-microsoft-365-agents-toolkit";

function readAtkExtensionBlock(manifest: OpenPluginManifest): AtkExtensionBlock | undefined {
  const raw = (manifest as unknown as Record<string, unknown>)[ATK_EXTENSION_KEY];
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return undefined;
  }
  return raw as AtkExtensionBlock;
}
