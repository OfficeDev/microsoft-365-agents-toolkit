// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { Result, err, ok } from "neverthrow";
import { DeclarativeLocator, TemplateFileEntry } from "../model/dataModel";
import { LoadedPackage } from "./packageDir";

/**
 * The zip-bytes declarative consume operation: open the channel `templates.zip`
 * and return one authored package's `v4/<kind>/<templateId>/` subtree — its
 * parsed `descriptor.json` / `pipeline.json` plus the raw `content/**` entries.
 *
 * This is the on-the-wire / bundled-floor sibling of `loadPackageDir` (which
 * reads the same `LoadedPackage` shape from a loose authoring directory): same
 * output noun, same deterministic order (INV-5), same Zip-Slip guard (INV-8) as
 * `openTemplatePackage`, but it locates by the `{kind, templateId}` package
 * prefix rather than the v3 `<language>/<scenario>/` mirror prefix. The two
 * readers are interchangeable, so `scaffold` runs identically from the shipped
 * floor and from the dev tree.
 *
 * Reads only the `v4/` declarative subtree; the zip's coexisting v3 mirror is
 * never read — v4 depends on no v3 content (INV-7). Shape validation is deferred
 * to `scaffold` (via `parseReplaceMap` / `parsePipeline`), keeping this a pure
 * "read what is there" step.
 *
 * Spec: docs/03-specs/operations/scaffolding/open-template-package.md (the
 * declarative-subtree variant of the same consume boundary).
 *
 * v4-owned (INV-7): imports no v3 symbol; v3 may call `openDeclarativePackage`,
 * but nothing here is tailored for v3.
 */

const SOURCE = "Scaffold";

/** The `v4/<kind>/<templateId>/` prefix this locator resolves to (trailing slash = boundary). */
function packageRoot(locator: DeclarativeLocator): string {
  return `v4/${locator.kind}/${locator.templateId}/`;
}

/**
 * Reject a stripped content path whose segments are empty / `.` / `..`
 * (Zip-Slip guard, INV-8). The renderer writes these paths to disk, so
 * containment is enforced at open time — same guard as `openTemplatePackage`.
 */
function isSafeRelativePath(rel: string): boolean {
  return rel.split("/").every((seg) => seg.length > 0 && seg !== "." && seg !== "..");
}

/** Parse one raw JSON entry, mapping a parse failure to `PackageFileInvalid`. */
function parseEntryJson(raw: string, file: string): Result<unknown, FxError> {
  try {
    const parsed: unknown = JSON.parse(raw);
    return ok(parsed);
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileInvalid",
        message: `The template package file "${file}" is not valid JSON.`,
      })
    );
  }
}

/** A package shape file that the package set is required to carry. */
function missingFile(file: string): FxError {
  return new SystemError({
    source: SOURCE,
    name: "PackageFileMissing",
    message: `The template package is missing "${file}".`,
  });
}

/**
 * Open the channel package and return the located declarative package. Pure
 * function of `(bytes, locator)`: no fs, no network, no render. Returns `Result`
 * per the toolkit-wide neverthrow rule rather than throwing; a corrupt archive,
 * a missing/invalid shape file, an empty content set, or a Zip-Slip path is a
 * `SystemError` (an authoring/packaging fault, not user-fixable).
 */
export function openDeclarativePackage(
  bytes: Buffer,
  locator: DeclarativeLocator
): Result<LoadedPackage, FxError> {
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

  const root = packageRoot(locator);
  const contentPrefix = `${root}content/`;
  let descriptorRaw: string | undefined;
  let pipelineRaw: string | undefined;
  const content: TemplateFileEntry[] = [];

  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) {
      continue; // INV-6 — directory entries excluded
    }
    const name = entry.entryName.replace(/\\/g, "/");
    if (!name.startsWith(root)) {
      continue; // INV-1 — trailing-slash prefix boundary; never reads the v3 mirror
    }
    if (name.startsWith(contentPrefix)) {
      const rel = name.slice(contentPrefix.length);
      if (!isSafeRelativePath(rel)) {
        return err(
          new SystemError({
            source: SOURCE,
            name: "TemplatePackageUnsafePath",
            message: `The resolved template package contains an unsafe entry path: "${entry.entryName}".`,
          })
        ); // INV-8 — Zip-Slip guard
      }
      content.push({ path: rel, data: entry.getData() });
      continue;
    }
    const rel = name.slice(root.length);
    if (rel === "descriptor.json") {
      descriptorRaw = entry.getData().toString("utf8");
    } else if (rel === "pipeline.json") {
      pipelineRaw = entry.getData().toString("utf8");
    }
    // Other root-level files (e.g. questions.json) are not part of LoadedPackage.
  }

  if (descriptorRaw === undefined) {
    return err(missingFile(`${root}descriptor.json`));
  }
  if (pipelineRaw === undefined) {
    return err(missingFile(`${root}pipeline.json`));
  }
  if (content.length === 0) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageContentMissing",
        message: `The template package "${root}" has no "content" entries.`,
      })
    );
  }

  const descriptor = parseEntryJson(descriptorRaw, `${root}descriptor.json`);
  if (descriptor.isErr()) {
    return err(descriptor.error);
  }
  const pipeline = parseEntryJson(pipelineRaw, `${root}pipeline.json`);
  if (pipeline.isErr()) {
    return err(pipeline.error);
  }

  content.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)); // INV-5 — deterministic order
  return ok({ descriptor: descriptor.value, pipeline: pipeline.value, content });
}
