// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { Result, err, ok } from "neverthrow";

/**
 * The v4 template-package consume operation: open resolved package bytes and
 * return one template's file entries, the locator prefix stripped, unrendered.
 *
 * Spec: docs/03-specs/operations/scaffolding/open-template-package.md
 * Upstream: docs/03-specs/operations/scaffolding/resolve-template-source.md
 *
 * This module is part of the v4 world. It imports no v3 symbol; v3 code may
 * call `openTemplatePackage`, but nothing here is tailored for v3.
 */

const SOURCE = "Scaffold";

/**
 * Names which template inside the package to open. The boundary
 * (open → locate → hand back entries) is permanent; this locator shape is
 * transitional — `{ language, scenario }` matches the package's current
 * `<language>/<scenario>/` layout and becomes `{ templateId }` when the
 * proposal §3 authoring layout ships. Only the resolved prefix changes; the
 * open/entry contract and every AC hold (INV-1).
 */
export interface TemplateLocator {
  language: string;
  scenario: string;
}

/** One file from the located template, rooted at the template's content. */
export interface TemplateFileEntry {
  /** Path relative to the located template's content root, forward-slash normalized. */
  path: string;
  /** The file's raw bytes, verbatim — unrendered, `.tpl` suffix intact. */
  data: Buffer;
}

/** The `<language>/<scenario>/` prefix this locator resolves to (trailing slash = boundary). */
function locatorPrefix(locator: TemplateLocator): string {
  return `${locator.language}/${locator.scenario}/`;
}

/**
 * Reject a stripped entry path whose segments are empty / `.` / `..` (Zip-Slip
 * guard, INV-8). A digest check proves the bytes equal what was published, not
 * that the published archive is free of traversal entries; the renderer writes
 * these paths to disk, so containment is enforced here.
 */
function isSafeRelativePath(rel: string): boolean {
  return rel.split("/").every((seg) => seg.length > 0 && seg !== "." && seg !== "..");
}

/**
 * Open the resolved template package and return the located template's file
 * entries, the locator prefix stripped. Pure function of `(bytes, locator)`:
 * no fs, no network, no render. Returns `Result` per the toolkit-wide
 * neverthrow rule rather than throwing.
 */
export function openTemplatePackage(
  bytes: Buffer,
  locator: TemplateLocator
): Result<TemplateFileEntry[], FxError> {
  let zip: AdmZip;
  try {
    zip = new AdmZip(bytes);
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "TemplatePackageCorrupt",
        message: "The resolved template package is not a valid archive.",
      })
    );
  }

  const prefix = locatorPrefix(locator);
  const entries: TemplateFileEntry[] = [];
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) {
      continue; // INV-6 — directory entries excluded
    }
    const name = entry.entryName.replace(/\\/g, "/");
    if (!name.startsWith(prefix)) {
      continue; // INV-1 — trailing-slash prefix boundary; "da/" never matches "da-basic/"
    }
    const rel = name.slice(prefix.length);
    if (!isSafeRelativePath(rel)) {
      return err(
        new SystemError({
          source: SOURCE,
          name: "TemplatePackageUnsafePath",
          message: `The resolved template package contains an unsafe entry path: "${entry.entryName}".`,
        })
      ); // INV-8 — Zip-Slip guard: reject empty / "." / ".." path segments
    }
    entries.push({ path: rel, data: entry.getData() });
  }

  if (entries.length === 0) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "TemplateNotFoundInPackage",
        message: `Template "${prefix}" was not found in the resolved package.`,
      })
    );
  }

  entries.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)); // INV-5 — deterministic order
  return ok(entries);
}
