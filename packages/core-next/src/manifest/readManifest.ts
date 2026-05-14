// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { ok, err, Result } from "neverthrow";
import { TeamsManifestWrapper } from "@microsoft/app-manifest";
import type { AtkContext } from "../core/context";
import { userError, systemError, AtkError } from "../core/error";
import { ManifestType } from "./types";
import { resolveManifest } from "./resolve";

// ─── Manifest path resolution ────────────────────────────────

/**
 * Default subfolder names where the toolkit stores manifest files, in priority order.
 * `.generated` is produced by the scaffold pipeline and takes precedence.
 */
const MANIFEST_SEARCH_PATHS = [path.join("appPackage", ".generated"), "appPackage"] as const;

const MANIFEST_FILENAME = "manifest.json";

/**
 * Locate the Teams app manifest.json in common toolkit project layouts.
 *
 * Search order:
 * 1. `<projectPath>/appPackage/.generated/manifest.json`
 * 2. `<projectPath>/appPackage/manifest.json`
 * 3. `<projectPath>/manifest.json`
 *
 * @returns The absolute path to the discovered manifest.json, or an error.
 */
export function getManifestPath(projectPath: string): Result<string, AtkError> {
  for (const subdir of MANIFEST_SEARCH_PATHS) {
    const candidate = path.join(projectPath, subdir, MANIFEST_FILENAME);
    if (fs.existsSync(candidate)) {
      return ok(candidate);
    }
  }

  // Fall back to project root
  const rootCandidate = path.join(projectPath, MANIFEST_FILENAME);
  if (fs.existsSync(rootCandidate)) {
    return ok(rootCandidate);
  }

  return err(
    userError(
      "ManifestNotFound",
      `Cannot find ${MANIFEST_FILENAME} in project. Searched: ${MANIFEST_SEARCH_PATHS.map((s) => path.join(projectPath, s)).join(", ")}, ${rootCandidate}`,
      { source: "ManifestRead" }
    )
  );
}

// ─── Read manifest ───────────────────────────────────────────

/**
 * Read a Teams app manifest and return both the raw JSON string and
 * the parsed `TeamsManifestWrapper`. No env-var resolution is performed.
 *
 * @param manifestPath  Absolute path to manifest.json
 */
export async function readTeamsManifest(
  manifestPath: string
): Promise<Result<{ raw: string; wrapper: TeamsManifestWrapper }, AtkError>> {
  try {
    let raw = await fs.promises.readFile(manifestPath, "utf8");
    // Strip BOM
    if (raw.charCodeAt(0) === 0xfeff) {
      raw = raw.slice(1);
    }
    const wrapper = TeamsManifestWrapper.fromJSON(raw);
    return ok({ raw, wrapper });
  } catch (e) {
    return err(
      systemError("ReadManifestFailed", `Failed to read manifest at ${manifestPath}: ${e}`, {
        source: "ManifestRead",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Read a Teams app manifest, resolve all `${{VAR}}` placeholders and
 * `$[file()]` expressions, then return the resolved JSON string and
 * a parsed `TeamsManifestWrapper` created from the resolved content.
 *
 * This is the primary entry point for packaging workflows that need
 * a fully-resolved manifest.
 *
 * @param manifestPath  Absolute path to manifest.json
 * @param ctx           Context for logging, telemetry, env-var lookup
 * @param envs          Optional explicit env-var map (falls back to process.env)
 */
export async function readAndResolveTeamsManifest(
  manifestPath: string,
  ctx: AtkContext,
  envs?: Readonly<Record<string, string>>
): Promise<Result<{ resolved: string; wrapper: TeamsManifestWrapper }, AtkError>> {
  const readResult = await readTeamsManifest(manifestPath);
  if (readResult.isErr()) return err(readResult.error);

  const resolveResult = await resolveManifest(readResult.value.raw, ctx, {
    envs,
    manifestType: ManifestType.TeamsManifest,
    fromPath: manifestPath,
    strict: true,
  });
  if (resolveResult.isErr()) return err(resolveResult.error);

  try {
    const wrapper = TeamsManifestWrapper.fromJSON(resolveResult.value);
    return ok({ resolved: resolveResult.value, wrapper });
  } catch (e) {
    return err(
      systemError("ParseResolvedManifestFailed", `Resolved manifest is not valid JSON: ${e}`, {
        source: "ManifestRead",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}
