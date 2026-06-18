// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";
import { Result, err, ok } from "neverthrow";
import { TemplateFileEntry } from "../model/dataModel";

/**
 * The on-disk template-package consume operation: load a declarative package
 * straight from its authored directory (`descriptor.json` + `pipeline.json` +
 * `content/**`) and return its two shape files (parsed, unvalidated) plus the
 * raw content entries.
 *
 * This is the on-disk sibling of `openTemplatePackage` (which opens a v3-shaped
 * zip from resolved bytes): same output noun (`TemplateFileEntry[]`), same
 * deterministic order (INV-5), but it reads loose files from the authored /
 * bundled layout rather than stripping a `<language>/<scenario>/` zip prefix.
 * Shape validation is deliberately deferred — `scaffold` (via `parseReplaceMap`
 * / `parsePipeline`) owns it, so this stays a pure "read what is there" step.
 *
 * Spec: docs/03-specs/operations/scaffolding/open-template-package.md (the
 * on-disk authoring-layout variant of the same consume boundary).
 *
 * v4-owned (INV-7): imports no v3 symbol; v3 may call `loadPackageDir`, but
 * nothing here is tailored for v3.
 */

const SOURCE = "Scaffold";

/** A parsed declarative package: the two shape files (unvalidated) + raw content. */
export interface LoadedPackage {
  /** The package's parsed `descriptor.json` (its `replaceMap` drives the render vars). */
  descriptor: unknown;
  /** The package's parsed `pipeline.json`. */
  pipeline: unknown;
  /** The opened `content/**` entries (raw bytes, `.tpl` suffix intact, sorted). */
  content: TemplateFileEntry[];
}

/** Read + parse one top-level package JSON file (EAFP: read, don't stat first). */
function readPackageJson(dir: string, file: string): Result<unknown, FxError> {
  let raw: string;
  try {
    raw = fs.readFileSync(path.join(dir, file), "utf8");
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileMissing",
        message: `The template package is missing "${file}".`,
      })
    );
  }
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

/** Recurse `content/**` into raw, forward-slash-pathed entries rooted at `root`. */
function walkContent(root: string, dir: string, out: TemplateFileEntry[]): void {
  for (const name of fs.readdirSync(dir)) {
    const full = path.join(dir, name);
    if (fs.statSync(full).isDirectory()) {
      walkContent(root, full, out);
    } else {
      out.push({
        path: path.relative(root, full).replace(/\\/g, "/"),
        data: fs.readFileSync(full),
      });
    }
  }
}

/** Load `content/**` as deterministically ordered raw entries (INV-5). */
function loadContent(contentRoot: string): Result<TemplateFileEntry[], FxError> {
  const entries: TemplateFileEntry[] = [];
  try {
    walkContent(contentRoot, contentRoot, entries);
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageContentMissing",
        message: `The template package's "content" directory could not be read.`,
      })
    );
  }
  entries.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
  return ok(entries);
}

/**
 * Load a declarative template package from its authored directory. Returns
 * `Result` per the toolkit-wide neverthrow rule; a missing/invalid shape file or
 * an unreadable `content` directory is a `SystemError` (an authoring/packaging
 * fault, not a user-fixable one).
 */
export function loadPackageDir(dir: string): Result<LoadedPackage, FxError> {
  const descriptor = readPackageJson(dir, "descriptor.json");
  if (descriptor.isErr()) {
    return err(descriptor.error);
  }
  const pipeline = readPackageJson(dir, "pipeline.json");
  if (pipeline.isErr()) {
    return err(pipeline.error);
  }
  const content = loadContent(path.join(dir, "content"));
  if (content.isErr()) {
    return err(content.error);
  }
  return ok({
    descriptor: descriptor.value,
    pipeline: pipeline.value,
    content: content.value,
  });
}
