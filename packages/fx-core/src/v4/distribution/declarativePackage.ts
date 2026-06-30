// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { Result, err, ok } from "neverthrow";
import { DeclarativeLocator, TemplateFileEntry } from "../model/dataModel";
import { LoadedPackage } from "./packageDir";

/** Open one declarative package subtree from channel zip bytes. See open-template-package spec. */

const SOURCE = "Scaffold";

/** The `v4/<kind>/<templateId>/` prefix this locator resolves to (trailing slash = boundary). */
function packageRoot(locator: DeclarativeLocator): string {
  return `v4/${locator.kind}/${locator.templateId}/`;
}

/** Zip-Slip guard for stripped content paths. */
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

/** Open the channel package and return the located declarative package. */
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
      continue;
    }
    const name = entry.entryName.replace(/\\/g, "/");
    if (!name.startsWith(root)) {
      continue;
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
        );
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
  const descriptor = parseEntryJson(descriptorRaw, `${root}descriptor.json`);
  if (descriptor.isErr()) {
    return err(descriptor.error);
  }
  const pipeline = parseEntryJson(pipelineRaw, `${root}pipeline.json`);
  if (pipeline.isErr()) {
    return err(pipeline.error);
  }

  content.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
  return ok({ descriptor: descriptor.value, pipeline: pipeline.value, content });
}
