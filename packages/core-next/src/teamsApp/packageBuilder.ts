// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import AdmZip from "adm-zip";
import semver from "semver";
import { ok, err, Result } from "neverthrow";
import {
  TeamsManifestWrapper,
  DeclarativeAgentManifestWrapper,
  CapabilityName,
  RuntimeType,
} from "@microsoft/app-manifest";
import type { AtkContext } from "../core/context";
import { userError, systemError, AtkError } from "../core/error";
import { ManifestType } from "../manifest/types";
import { resolveManifest, resolveEnvPlaceholders } from "../manifest/resolve";
import { readAndResolveTeamsManifest, getManifestPath } from "../manifest/readManifest";

// ─── Public API ──────────────────────────────────────────────

/**
 * Options for building a Teams app package (.zip).
 */
export interface PackageBuildOptions {
  /** Path to the manifest.json (absolute or relative to projectPath). */
  manifestPath?: string;
  /** Output path for the .zip file. */
  outputZipPath: string;
  /** Folder to write resolved manifest files (DA, plugin etc.) alongside the ZIP. */
  outputFolder?: string;
  /** Project root directory. */
  projectPath: string;
  /** Optional env-var map; falls back to process.env when omitted. */
  envs?: Readonly<Record<string, string>>;
}

/**
 * The result of a successful package build.
 */
export interface PackageBuildResult {
  /** Absolute path to the produced ZIP. */
  zipPath: string;
  /** Absolute path to the resolved manifest.json written to disk. */
  jsonPath: string;
  /** Number of files in the ZIP. */
  fileCount: number;
}

const SOURCE = "PackageBuilder";
const HELP_LINK = "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage";

/**
 * Build a Teams app package (.zip) from a manifest directory.
 *
 * This is the production-grade packaging pipeline ported from
 * fx-core's CreateAppPackageDriver.build().
 *
 * Pipeline:
 * 1. Locate manifest.json (.generated folder detection)
 * 2. Read + resolve env vars
 * 3. Validate & collect icons (color, outline, color32x32 for v1.21+)
 * 4. Process localization files
 * 5. Process API-based compose extensions (API spec + adaptive card templates)
 * 6. Process DA manifests → plugin actions → embedded knowledge
 * 7. Build ZIP with path-traversal protection
 * 8. Write resolved manifest JSON alongside ZIP
 */
export async function buildAppPackage(
  ctx: AtkContext,
  options: PackageBuildOptions
): Promise<Result<PackageBuildResult, AtkError>> {
  const { projectPath, outputZipPath, outputFolder, envs } = options;

  // ── 1. Locate manifest ──────────────────────────────────
  const generatedFolder = path.join(projectPath, "appPackage", ".generated");
  const hasGenerated =
    fs.existsSync(generatedFolder) && fs.existsSync(path.join(generatedFolder, "manifest.json"));

  let manifestPath: string;
  if (options.manifestPath) {
    manifestPath = path.isAbsolute(options.manifestPath)
      ? options.manifestPath
      : path.join(projectPath, options.manifestPath);
  } else if (hasGenerated) {
    manifestPath = path.join(generatedFolder, "manifest.json");
  } else {
    const found = getManifestPath(projectPath);
    if (found.isErr()) return err(found.error);
    manifestPath = found.value;
  }

  // ── 2. Read + resolve manifest ──────────────────────────
  const readRes = await readAndResolveTeamsManifest(manifestPath, ctx, envs);
  if (readRes.isErr()) return err(readRes.error);
  const { resolved: resolvedManifestJson, wrapper: manifest } = readRes.value;

  const appDirectory = path.dirname(hasGenerated ? generatedFolder : manifestPath);

  // ── 3. Icons ────────────────────────────────────────────
  const iconFiles = [manifest.icons.color, manifest.icons.outline];
  const manifestVersion =
    manifest.manifestVersion === "devPreview"
      ? semver.coerce("1.19.0")
      : semver.coerce(manifest.manifestVersion);

  if (manifestVersion && semver.gte(manifestVersion, "1.21.0")) {
    // v1.21+ supports 32×32 color icon
    const icons = manifest.icons as { color: string; outline: string; color32x32?: string };
    if (icons.color32x32) {
      iconFiles.push(icons.color32x32);
    }
  }

  for (const icon of iconFiles) {
    const iconAbsPath = path.resolve(appDirectory, icon);
    const checkRes = validateReferencedFile(iconAbsPath, appDirectory);
    if (checkRes.isErr()) return err(checkRes.error);
  }

  // ── 4. Localization pre-check ───────────────────────────
  const additionalLanguages = getAdditionalLanguages(manifest, manifestVersion);
  const defaultLanguageFile = getDefaultLanguageFile(manifest, manifestVersion);

  for (const lang of additionalLanguages) {
    const langPath = path.join(appDirectory, lang.file);
    if (!fs.existsSync(langPath)) {
      return err(fileNotFound(langPath));
    }
  }
  if (defaultLanguageFile) {
    const defPath = path.join(appDirectory, defaultLanguageFile);
    if (!fs.existsSync(defPath)) {
      return err(fileNotFound(defPath));
    }
  }

  // ── 5–8. Build ZIP ─────────────────────────────────────
  const zip = new AdmZip();

  // Add resolved manifest.json
  zip.addFile("manifest.json", Buffer.from(resolvedManifestJson));

  // Add icons
  for (const icon of iconFiles) {
    const dir = path.dirname(icon);
    zip.addLocalFile(path.resolve(appDirectory, icon), dir === "." ? "" : dir);
  }

  // Add localization files
  for (const lang of additionalLanguages) {
    const langAbsPath = path.resolve(appDirectory, lang.file);
    const relCheck = validateReferencedFile(langAbsPath, appDirectory);
    if (relCheck.isErr()) return err(relCheck.error);

    const resolvedLoc = await resolveFileContent(langAbsPath, envs);
    if (resolvedLoc.isErr()) return err(resolvedLoc.error);

    const rel = path.relative(appDirectory, langAbsPath);
    zip.addFile(normalizeZipPath(rel), Buffer.from(resolvedLoc.value));
  }
  if (defaultLanguageFile) {
    const defAbsPath = path.resolve(appDirectory, defaultLanguageFile);
    const relCheck = validateReferencedFile(defAbsPath, appDirectory);
    if (relCheck.isErr()) return err(relCheck.error);

    const resolvedDef = await resolveFileContent(defAbsPath, envs);
    if (resolvedDef.isErr()) return err(resolvedDef.error);

    const rel = path.relative(appDirectory, defAbsPath);
    zip.addFile(normalizeZipPath(rel), Buffer.from(resolvedDef.value));
  }

  // API-based compose extensions
  const ceRes = await processComposeExtensions(
    zip,
    manifest,
    manifestVersion,
    appDirectory,
    ctx,
    envs
  );
  if (ceRes.isErr()) return err(ceRes.error);

  // Declarative Agents
  const daRes = await processDeclarativeAgents(
    zip,
    manifest,
    manifestVersion,
    appDirectory,
    hasGenerated ? generatedFolder : appDirectory,
    ctx,
    envs,
    outputFolder
  );
  if (daRes.isErr()) return err(daRes.error);

  // ── Write ZIP ───────────────────────────────────────────
  const zipAbsPath = path.isAbsolute(outputZipPath)
    ? outputZipPath
    : path.join(projectPath, outputZipPath);
  await fs.promises.mkdir(path.dirname(zipAbsPath), { recursive: true });
  zip.writeZip(zipAbsPath);

  // ── Write resolved manifest JSON ────────────────────────
  const jsonDir = outputFolder
    ? path.isAbsolute(outputFolder)
      ? outputFolder
      : path.join(projectPath, outputFolder)
    : path.dirname(zipAbsPath);
  await fs.promises.mkdir(jsonDir, { recursive: true });

  const envName = envs?.["TEAMSFX_ENV"] ?? process.env.TEAMSFX_ENV ?? "local";
  const jsonFileName = path.join(jsonDir, `manifest.${envName}.json`);
  await writeReadOnlyFile(jsonFileName, resolvedManifestJson);

  ctx.logger.info(`App package built: ${zipAbsPath}`);

  return ok({
    zipPath: zipAbsPath,
    jsonPath: jsonFileName,
    fileCount: zip.getEntries().length,
  });
}

// ─── Compose Extensions ──────────────────────────────────────

async function processComposeExtensions(
  zip: AdmZip,
  manifest: TeamsManifestWrapper,
  manifestVersion: semver.SemVer | null,
  appDirectory: string,
  ctx: AtkContext,
  envs?: Readonly<Record<string, string>>
): Promise<Result<void, AtkError>> {
  if (!manifestVersion || !semver.gte(manifestVersion, "1.17.0")) return ok(undefined);

  const extensions = manifest.composeExtensions;
  if (extensions.length === 0) return ok(undefined);

  const ext = extensions[0] as Record<string, unknown>;
  const ceType = ext.composeExtensionType as string | undefined;
  const apiSpecFile = ext.apiSpecificationFile as string | undefined;
  const commands = ext.commands as Array<Record<string, unknown>> | undefined;

  if (ceType !== "apiBased" || !apiSpecFile) return ok(undefined);

  // Add API spec
  const specAbsPath = path.resolve(appDirectory, apiSpecFile);
  const checkRes = validateReferencedFile(specAbsPath, appDirectory);
  if (checkRes.isErr()) return err(checkRes.error);

  const resolvedSpec = await addFileWithEnvResolution(
    zip,
    apiSpecFile,
    specAbsPath,
    ManifestType.ApiSpec,
    ctx,
    envs
  );
  if (resolvedSpec.isErr()) return err(resolvedSpec.error);

  // Add adaptive card templates for commands
  if (commands) {
    for (const cmd of commands) {
      const cardFile = cmd.apiResponseRenderingTemplateFile as string | undefined;
      if (!cardFile) continue;
      const cardAbsPath = path.resolve(appDirectory, cardFile);
      const cardCheck = validateReferencedFile(cardAbsPath, appDirectory);
      if (cardCheck.isErr()) return err(cardCheck.error);
      const dir = path.dirname(cardFile);
      zip.addLocalFile(cardAbsPath, dir === "." ? "" : dir);
    }
  }

  return ok(undefined);
}

// ─── Declarative Agents ──────────────────────────────────────

async function processDeclarativeAgents(
  zip: AdmZip,
  manifest: TeamsManifestWrapper,
  manifestVersion: semver.SemVer | null,
  appDirectory: string,
  daBaseDir: string,
  ctx: AtkContext,
  envs?: Readonly<Record<string, string>>,
  outputFolder?: string
): Promise<Result<void, AtkError>> {
  if (!manifestVersion || !semver.gte(manifestVersion, "1.19.0")) return ok(undefined);

  const daRefs = manifest.declarativeAgents;
  if (daRefs.length === 0) return ok(undefined);

  const daRef = daRefs[0];
  const daFile = (daRef as unknown as Record<string, unknown>).file as string | undefined;
  if (!daFile) return ok(undefined);

  const daAbsPath = path.resolve(daBaseDir, daFile);
  const daCheck = validateReferencedFile(daAbsPath, appDirectory);
  if (daCheck.isErr()) return err(daCheck.error);

  // Add DA manifest with env-var resolution
  const addDaRes = await addFileWithEnvResolution(
    zip,
    daFile,
    daAbsPath,
    ManifestType.DeclarativeCopilotManifest,
    ctx,
    envs
  );
  if (addDaRes.isErr()) return err(addDaRes.error);

  // Write resolved DA manifest to output folder if requested
  if (outputFolder) {
    const daOutputPath = path.join(outputFolder, path.relative(appDirectory, daAbsPath));
    await writeReadOnlyFile(daOutputPath, addDaRes.value);
  }

  // Parse DA manifest for actions and capabilities
  let daWrapper: DeclarativeAgentManifestWrapper;
  try {
    daWrapper = DeclarativeAgentManifestWrapper.fromJSON(addDaRes.value);
  } catch (e) {
    return err(
      systemError("ParseDAManifestFailed", `Failed to parse DA manifest: ${e}`, {
        source: SOURCE,
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }

  // Process plugin actions
  for (const action of daWrapper.actions) {
    const pluginFile = (action as Record<string, unknown>).file as string | undefined;
    if (!pluginFile) continue;

    const pluginAbsPath = path.resolve(path.dirname(daAbsPath), pluginFile);
    const pluginRelPath = path.relative(daBaseDir, pluginAbsPath);
    const useForwardSlash = daFile.includes("/") || pluginFile.includes("/");

    const pluginRes = await processPlugin(
      zip,
      normalizePathSep(pluginRelPath, useForwardSlash),
      daBaseDir,
      appDirectory,
      ctx,
      envs,
      outputFolder
    );
    if (pluginRes.isErr()) return err(pluginRes.error);
  }

  // Process embedded knowledge files
  const ekRes = processEmbeddedKnowledge(zip, daWrapper, appDirectory);
  if (ekRes.isErr()) return err(ekRes.error);

  return ok(undefined);
}

// ─── Plugin Processing ───────────────────────────────────────

async function processPlugin(
  zip: AdmZip,
  pluginRelativePath: string,
  baseDir: string,
  appDirectory: string,
  ctx: AtkContext,
  envs?: Readonly<Record<string, string>>,
  outputFolder?: string
): Promise<Result<void, AtkError>> {
  const pluginAbsPath = path.resolve(baseDir, pluginRelativePath);
  const checkRes = validateReferencedFile(pluginAbsPath, baseDir);
  if (checkRes.isErr()) return err(checkRes.error);

  // Read and parse plugin manifest
  let pluginContent: Record<string, unknown>;
  try {
    const raw = await fs.promises.readFile(pluginAbsPath, "utf8");
    pluginContent = JSON.parse(raw);
  } catch (e) {
    return err(
      userError(
        "InvalidPluginManifest",
        `Failed to parse plugin manifest at ${pluginAbsPath}: ${e}`,
        { source: SOURCE, inner: e instanceof Error ? e : new Error(String(e)) }
      )
    );
  }

  // Inline adaptive card templates from external files
  let hadExternalCards = false;
  const functions = pluginContent.functions as Array<Record<string, unknown>> | undefined;
  if (functions) {
    for (const func of functions) {
      const staticTemplateFile = getStaticTemplateFile(func, pluginAbsPath, appDirectory);
      if (!staticTemplateFile) continue;

      try {
        const cardContent = await fs.promises.readFile(staticTemplateFile, "utf8");
        const cardJson = JSON.parse(cardContent);
        const caps = func.capabilities as Record<string, unknown> | undefined;
        const semantics = caps?.response_semantics as Record<string, unknown> | undefined;
        if (semantics?.static_template) {
          semantics.static_template = cardJson;
          hadExternalCards = true;
        }
      } catch (e) {
        ctx.logger.warning(
          `Failed to inline adaptive card template from ${staticTemplateFile}: ${e}`
        );
      }
    }
  }

  // Strip underscores from namespace (Teams requirement)
  const ns = pluginContent.namespace as string | undefined;
  let hadNamespaceFix = false;
  if (ns?.includes("_")) {
    pluginContent.namespace = ns.replace(/_/g, "");
    hadNamespaceFix = true;
  }

  // If modifications were made, serialize the modified content and resolve
  let pluginJsonStr: string;
  if (hadExternalCards || hadNamespaceFix) {
    pluginJsonStr = JSON.stringify(pluginContent, null, 4);
    const funcRes = await resolveManifest(pluginJsonStr, ctx, {
      envs,
      manifestType: ManifestType.PluginManifest,
      fromPath: pluginAbsPath,
      strict: true,
    });
    if (funcRes.isErr()) return err(funcRes.error);
    pluginJsonStr = funcRes.value;

    zip.addFile(pluginRelativePath, Buffer.from(pluginJsonStr));
  } else {
    const addRes = await addFileWithEnvResolution(
      zip,
      pluginRelativePath,
      pluginAbsPath,
      ManifestType.PluginManifest,
      ctx,
      envs
    );
    if (addRes.isErr()) return err(addRes.error);
    pluginJsonStr = addRes.value;
  }

  if (outputFolder) {
    const outPath = path.join(outputFolder, path.relative(appDirectory, pluginAbsPath));
    await writeReadOnlyFile(outPath, pluginJsonStr);
  }

  // Process runtimes (OpenAPI specs, MCP tool descriptions)
  return processPluginRuntimes(zip, pluginRelativePath, baseDir, ctx, envs);
}

async function processPluginRuntimes(
  zip: AdmZip,
  pluginRelativePath: string,
  baseDir: string,
  ctx: AtkContext,
  envs?: Readonly<Record<string, string>>
): Promise<Result<void, AtkError>> {
  const pluginAbsPath = path.resolve(baseDir, pluginRelativePath);
  let pluginData: Record<string, unknown>;
  try {
    const content = await fs.promises.readFile(pluginAbsPath, "utf8");
    pluginData = JSON.parse(content);
  } catch {
    return ok(undefined);
  }

  const runtimes = pluginData.runtimes as Array<Record<string, unknown>> | undefined;
  if (!runtimes) return ok(undefined);

  for (const runtime of runtimes) {
    const type = runtime.type as string | undefined;
    const spec = runtime.spec as Record<string, unknown> | undefined;

    if (type === RuntimeType.OpenApi && spec?.url) {
      const specFile = path.resolve(path.dirname(pluginAbsPath), spec.url as string);
      const checkRes = validateReferencedFile(specFile, baseDir);
      if (checkRes.isErr()) return err(checkRes.error);

      const entryName = path.relative(baseDir, specFile);
      const useForwardSlash =
        pluginRelativePath.includes("/") || (spec.url as string).includes("/");

      const addRes = await addFileWithEnvResolution(
        zip,
        normalizePathSep(entryName, useForwardSlash),
        specFile,
        ManifestType.ApiSpec,
        ctx,
        envs
      );
      if (addRes.isErr()) return err(addRes.error);
    } else if (type === "RemoteMCPServer") {
      const mcpDesc = spec?.mcp_tool_description as Record<string, unknown> | undefined;
      const mcpFile = mcpDesc?.file as string | undefined;
      if (!mcpFile) continue;

      const mcpAbsPath = path.resolve(path.dirname(pluginAbsPath), mcpFile);
      const checkRes = validateReferencedFile(mcpAbsPath, baseDir);
      if (checkRes.isErr()) return err(checkRes.error);

      const entryName = path.relative(baseDir, mcpAbsPath);
      const dir = path.dirname(entryName);
      zip.addLocalFile(mcpAbsPath, dir === "." ? "" : dir);
    }
  }

  return ok(undefined);
}

// ─── Embedded Knowledge ──────────────────────────────────────

function processEmbeddedKnowledge(
  zip: AdmZip,
  daWrapper: DeclarativeAgentManifestWrapper,
  appDirectory: string
): Result<void, AtkError> {
  const ekCapabilities = daWrapper.capabilities.filter(
    (cap) => (cap as Record<string, unknown>).name === CapabilityName.EmbeddedKnowledge
  );
  if (ekCapabilities.length === 0) return ok(undefined);

  const fileSet = new Set<string>();
  for (const cap of ekCapabilities) {
    const files = (cap as Record<string, unknown>).files as
      | Array<Record<string, unknown>>
      | undefined;
    if (!files) continue;
    for (const f of files) {
      const filePath = f.file as string | undefined;
      if (filePath) fileSet.add(filePath);
    }
  }

  for (const file of fileSet) {
    const absPath = path.resolve(appDirectory, file);
    const checkRes = validateReferencedFile(absPath, appDirectory);
    if (checkRes.isErr()) return err(checkRes.error);

    const dir = path.dirname(file);
    zip.addLocalFile(absPath, dir === "." ? "" : dir);
  }

  return ok(undefined);
}

// ─── Helpers ─────────────────────────────────────────────────

function validateReferencedFile(absPath: string, baseDir: string): Result<void, AtkError> {
  if (!fs.existsSync(absPath)) {
    return err(fileNotFound(absPath));
  }
  const rel = path.relative(baseDir, absPath);
  if (rel.startsWith("..")) {
    return err(
      userError(
        "FileOutsideDirectory",
        `Referenced file "${absPath}" is outside the app package directory "${baseDir}".`,
        { source: SOURCE, help: HELP_LINK }
      )
    );
  }
  return ok(undefined);
}

function fileNotFound(filePath: string): AtkError {
  return userError("FileNotFound", `Referenced file not found: ${filePath}`, {
    source: SOURCE,
    help: HELP_LINK,
  });
}

async function addFileWithEnvResolution(
  zip: AdmZip,
  entryName: string,
  filePath: string,
  manifestType: ManifestType,
  ctx: AtkContext,
  envs?: Readonly<Record<string, string>>
): Promise<Result<string, AtkError>> {
  let content: string;
  try {
    content = await fs.promises.readFile(filePath, "utf8");
  } catch (e) {
    return err(
      systemError("ReadFileFailed", `Failed to read ${filePath}: ${e}`, {
        source: SOURCE,
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }

  const resolved = await resolveManifest(content, ctx, {
    envs,
    manifestType,
    fromPath: filePath,
    strict: manifestType !== ManifestType.ApiSpec,
  });
  if (resolved.isErr()) return err(resolved.error);

  zip.addFile(entryName, Buffer.from(resolved.value));
  return ok(resolved.value);
}

async function resolveFileContent(
  filePath: string,
  envs?: Readonly<Record<string, string>>
): Promise<Result<string, AtkError>> {
  try {
    const raw = await fs.promises.readFile(filePath, "utf8");
    const { content } = resolveEnvPlaceholders(raw, envs);
    return ok(content);
  } catch (e) {
    return err(
      systemError("ReadFileFailed", `Failed to read ${filePath}: ${e}`, {
        source: SOURCE,
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

async function writeReadOnlyFile(filePath: string, content: string): Promise<void> {
  await fs.promises.mkdir(path.dirname(filePath), { recursive: true });
  try {
    await fs.promises.chmod(filePath, 0o666);
  } catch {
    // File may not exist yet
  }
  await fs.promises.writeFile(filePath, content);
  await fs.promises.chmod(filePath, 0o444);
}

function normalizeZipPath(p: string): string {
  return p.replace(/\\/g, "/");
}

function normalizePathSep(p: string, useForwardSlash: boolean): string {
  return useForwardSlash ? p.replace(/\\/g, "/") : p;
}

function getStaticTemplateFile(
  func: Record<string, unknown>,
  pluginPath: string,
  appDirectory: string
): string | undefined {
  const caps = func.capabilities as Record<string, unknown> | undefined;
  const semantics = caps?.response_semantics as Record<string, unknown> | undefined;
  const staticTemplate = semantics?.static_template as Record<string, unknown> | undefined;
  const templateFile = staticTemplate?.file as string | undefined;
  if (!templateFile) return undefined;

  const candidate = path.resolve(appDirectory, templateFile);
  if (fs.existsSync(candidate)) return candidate;

  const fallback = path.resolve(path.dirname(pluginPath), templateFile);
  if (fs.existsSync(fallback)) return fallback;

  return undefined;
}

function getAdditionalLanguages(
  manifest: TeamsManifestWrapper,
  manifestVersion: semver.SemVer | null
): Array<{ file: string }> {
  if (!manifestVersion || !semver.gte(manifestVersion, "1.5.0")) return [];
  const data = manifest as unknown as Record<string, unknown>;
  const locInfo = data.localizationInfo as Record<string, unknown> | undefined;
  const langs = locInfo?.additionalLanguages as Array<{ file: string }> | undefined;
  return langs ?? [];
}

function getDefaultLanguageFile(
  manifest: TeamsManifestWrapper,
  manifestVersion: semver.SemVer | null
): string | undefined {
  if (!manifestVersion || !semver.gte(manifestVersion, "1.19.0")) return undefined;
  const data = manifest as unknown as Record<string, unknown>;
  const locInfo = data.localizationInfo as Record<string, unknown> | undefined;
  return locInfo?.defaultLanguageFile as string | undefined;
}
