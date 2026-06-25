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

/** Load the bundled-floor create selector. See resolve-build-target spec. */

const SOURCE = "Scaffold";

type SelectorKind = "create" | "modify";

/** The selector's fixed entry path inside the channel `templates.zip`. */
function selectorEntry(kind: SelectorKind): string {
  return `v4/${kind}/selector.json`;
}

/** Read and JSON-parse the single selector entry shared by both projections. */
function readSelectorRaw(bytes: Buffer, kind: SelectorKind): Result<unknown, FxError> {
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

  const selectorPath = selectorEntry(kind);
  let selectorRaw: string | undefined;
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) {
      continue;
    }
    const name = entry.entryName.replace(/\\/g, "/");
    if (name === selectorPath) {
      selectorRaw = entry.getData().toString("utf8");
      break;
    }
  }

  if (selectorRaw === undefined) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileMissing",
        message: `The template package is missing "${selectorPath}".`,
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
        message: `The template package file "${selectorPath}" is not valid JSON.`,
      })
    );
  }

  return ok(parsed);
}

export function openSelector(bytes: Buffer, kind: SelectorKind): Result<SelectorSpec, FxError> {
  const raw = readSelectorRaw(bytes, kind);
  if (raw.isErr()) {
    return err(raw.error);
  }
  return parseSelectorSpec(raw.value);
}

export function openSelectorPresentation(
  bytes: Buffer,
  kind: SelectorKind
): Result<SelectorPresentation, FxError> {
  const raw = readSelectorRaw(bytes, kind);
  if (raw.isErr()) {
    return err(raw.error);
  }
  return parseSelectorPresentation(raw.value);
}

export function openCreateSelector(bytes: Buffer): Result<SelectorSpec, FxError> {
  return openSelector(bytes, "create");
}

export function openModifySelector(bytes: Buffer): Result<SelectorSpec, FxError> {
  return openSelector(bytes, "modify");
}

/** Open the selector presentation projection for the live Q1 prompt face. */
export function openCreateSelectorPresentation(
  bytes: Buffer
): Result<SelectorPresentation, FxError> {
  return openSelectorPresentation(bytes, "create");
}

export function openModifySelectorPresentation(
  bytes: Buffer
): Result<SelectorPresentation, FxError> {
  return openSelectorPresentation(bytes, "modify");
}
