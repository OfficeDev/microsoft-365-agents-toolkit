// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { Result, err, ok } from "neverthrow";
import { SelectorSpec } from "../buildTarget/resolveBuildTarget";
import {
  SelectorPresentation,
  parseSelectorPresentation,
  parseSelectorSpec,
} from "../buildTarget/parseSelector";

/**
 * The bundled-floor load face of the create `selector.json`
 * (resolve-build-target AC-22). Open the channel `templates.zip` bytes, read
 * the single `v4/create/selector.json` entry, and project it onto a
 * `SelectorSpec` via `parseSelectorSpec`.
 *
 * The sibling of `openDeclarativePackage` for the routing table: same zip-bytes
 * input, same archive read, but it locates the one selector entry rather than a
 * `{ kind, templateId }` package subtree. It lets a surface (today the v3
 * declarative-agent bridge, `route-declarative-via-selector`) route through the
 * shipped selector without a hand-coded template-id table.
 *
 * The zip-read faults — a non-archive byte string, a floor missing the selector
 * entry, or a non-JSON entry — are `SystemError`s (packaging faults, not
 * user-fixable). Structural validity stays `parseSelectorSpec`'s contract
 * (resolve-build-target AC-20).
 *
 * v4-owned (INV-7): imports no v3 symbol; v3 may call `openCreateSelector`, but
 * nothing here is tailored for v3.
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-build-target.md (AC-22).
 */

const SOURCE = "Scaffold";

/** The create selector's fixed entry path inside the channel `templates.zip`. */
const SELECTOR_ENTRY = "v4/create/selector.json";

/**
 * Read + JSON-parse the single `selector.json` entry — the shared zip-read both
 * the routing (`openCreateSelector`) and presentation
 * (`openCreateSelectorPresentation`) openers compose, so the archive is opened
 * and located one way. The zip-read faults (a non-archive byte string, a floor
 * missing the entry, a non-JSON entry) are `SystemError`s (packaging faults);
 * structural validity stays the two parsers' contract.
 */
function readSelectorRaw(bytes: Buffer): Result<unknown, FxError> {
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

  let selectorRaw: string | undefined;
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) {
      continue;
    }
    const name = entry.entryName.replace(/\\/g, "/");
    if (name === SELECTOR_ENTRY) {
      selectorRaw = entry.getData().toString("utf8");
      break;
    }
  }

  if (selectorRaw === undefined) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileMissing",
        message: `The template package is missing "${SELECTOR_ENTRY}".`,
      })
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(selectorRaw);
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileInvalid",
        message: `The template package file "${SELECTOR_ENTRY}" is not valid JSON.`,
      })
    );
  }

  return ok(parsed);
}

export function openCreateSelector(bytes: Buffer): Result<SelectorSpec, FxError> {
  const raw = readSelectorRaw(bytes);
  if (raw.isErr()) {
    return err(raw.error);
  }
  return parseSelectorSpec(raw.value);
}

/**
 * The presentation sibling of `openCreateSelector`: read the same
 * `v4/create/selector.json` entry, but project it onto the `SelectorPresentation`
 * the live Q1 prompt face renders (each question's `title` / `placeholder` /
 * `staticOptions`) rather than the routing `SelectorSpec`.
 *
 * Spec: docs/03-specs/operations/scaffolding/walk-create-selector.md (WCS-07).
 */
export function openCreateSelectorPresentation(
  bytes: Buffer
): Result<SelectorPresentation, FxError> {
  const raw = readSelectorRaw(bytes);
  if (raw.isErr()) {
    return err(raw.error);
  }
  return parseSelectorPresentation(raw.value);
}
