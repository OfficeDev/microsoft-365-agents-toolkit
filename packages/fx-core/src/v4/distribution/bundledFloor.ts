// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";
import { getTemplatesFolder } from "../../folder";
import { BundledFloor, computeDigest } from "./templateSource";

/**
 * The bundled floor: the v4 template package baked into the engine so the
 * scaffold never depends on the network being reachable (offline-by-default).
 *
 * The build step drops two files under `<floor dir>`:
 *   - `floor.json`    → `{ "version": "<semver>" }` (the baked version)
 *   - `templates.zip` → the package bytes
 *
 * The digest is NOT baked: it is computed from the bytes at load time so
 * `computeDigest` stays the single authority (spec decision #6 / INV-3).
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-template-source.md
 */

const SOURCE = "Scaffold";

/** The directory the build bakes the floor into. */
export function bundledFloorDir(): string {
  return path.join(getTemplatesFolder(), "v4");
}

/** Build a {@link BundledFloor} from raw bytes — digest computed, never baked. */
export function bundledFloorFrom(version: string, bytes: Buffer, location: string): BundledFloor {
  return { version, digest: computeDigest(bytes), location };
}

interface FloorManifest {
  version: string;
}

function isFloorManifest(value: unknown): value is FloorManifest {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as Record<string, unknown>).version === "string"
  );
}

/**
 * Load the floor baked under `floorDir` (defaults to {@link bundledFloorDir}).
 * Missing/malformed bake artifacts are a hard error — a build without a floor
 * cannot scaffold offline, so we refuse rather than silently degrade.
 */
export function loadBundledFloor(floorDir: string = bundledFloorDir()): BundledFloor {
  const manifestPath = path.join(floorDir, "floor.json");
  const zipPath = path.join(floorDir, "templates.zip");

  let manifest: unknown;
  try {
    manifest = fs.readJsonSync(manifestPath);
  } catch {
    throw new SystemError({
      source: SOURCE,
      name: "BundledFloorMissing",
      message: `The bundled floor manifest is missing or unreadable at "${manifestPath}".`,
    });
  }
  if (!isFloorManifest(manifest)) {
    throw new SystemError({
      source: SOURCE,
      name: "BundledFloorMalformed",
      message: `The bundled floor manifest at "${manifestPath}" has no string "version".`,
    });
  }

  let bytes: Buffer;
  try {
    bytes = fs.readFileSync(zipPath);
  } catch {
    throw new SystemError({
      source: SOURCE,
      name: "BundledFloorMissing",
      message: `The bundled floor package is missing or unreadable at "${zipPath}".`,
    });
  }

  return bundledFloorFrom(manifest.version, bytes, zipPath);
}
