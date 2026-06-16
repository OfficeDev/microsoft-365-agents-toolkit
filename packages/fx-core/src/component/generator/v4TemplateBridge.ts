// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import { merge } from "lodash";
import path from "path";
import { TelemetryProperty } from "../../common/telemetry";
import templateConfig from "../../common/templates-config.json";
import { TemplateFileEntry, TemplateLocator, TemplateSource } from "../../v4";
import * as bundledFloorMod from "../../v4/distribution/bundledFloor";
import * as templatePackageMod from "../../v4/distribution/templatePackage";
import * as templateSourceMod from "../../v4/distribution/templateSource";
import * as templateSourcePortMod from "../../v4/distribution/templateSourcePort";
import { defaultTryLimits } from "./constant";
import { TemplateOutputPathError } from "./error";
import { GeneratorContext } from "./generatorAction";

export const v4TemplateBridgeDeps = {
  createTemplateSourcePort: templateSourcePortMod.createTemplateSourcePort,
  loadBundledFloor: bundledFloorMod.loadBundledFloor,
  resolveTemplateSource: templateSourceMod.resolveTemplateSource,
  loadResolvedPackage: templateSourcePortMod.loadResolvedPackage,
  openTemplatePackage: templatePackageMod.openTemplatePackage,
};

/**
 * Resolve a template entry's relative name to an absolute path and verify it
 * stays within `destination`. The entry name originates from the (untrusted)
 * template archive, so a `../` segment could otherwise escape the project
 * directory (zip-slip). Throws `TemplateOutputPathError` on escape.
 */
function resolveTemplateOutputPath(destination: string, entryName: string): string {
  const base = path.resolve(destination);
  const outputPath = path.resolve(base, entryName);
  const relative = path.relative(base, outputPath);
  // Reject only an actual parent-directory escape: a relative path that is the
  // `..` segment itself or starts with `..<sep>`, or an absolute path. A leading
  // `""` means the entry resolves to `base` itself (no filename). A filename
  // that merely starts with ".." (e.g. "..foo") stays in-root and is allowed.
  if (
    relative === "" ||
    relative === ".." ||
    relative.startsWith(".." + path.sep) ||
    path.isAbsolute(relative)
  ) {
    throw new TemplateOutputPathError(entryName);
  }
  return outputPath;
}

/**
 * Render the located template's entries onto disk using the v3
 * `GeneratorContext` rename/data/filter functions verbatim, so the output is
 * byte-identical to the legacy `unzip` path.
 *
 * `openTemplatePackage` strips the `<language>/<scenario>/` prefix, but the v3
 * `filterFn`/`fileNameReplaceFn` expect entry paths still rooted at
 * `${context.name}/`; the prefix is re-added here before applying them. The
 * data-replace function receives the basename, matching `unzip`'s
 * `dataReplaceFn(entry.name, ...)` call.
 */
export async function renderTemplateEntries(
  context: GeneratorContext,
  entries: TemplateFileEntry[]
): Promise<string[]> {
  const output: string[] = [];
  for (const entry of entries) {
    const entryName = `${context.name}/${entry.path}`;
    if (context.filterFn && !context.filterFn(entryName)) {
      continue;
    }
    const finalName = context.fileNameReplaceFn
      ? context.fileNameReplaceFn(entryName, entry.data)
      : entryName;
    const finalData = context.fileDataReplaceFn
      ? context.fileDataReplaceFn(path.basename(entry.path), entry.data)
      : entry.data;
    const filePath = resolveTemplateOutputPath(context.destination, finalName);
    await fs.ensureDir(path.dirname(filePath));
    await fs.writeFile(filePath, finalData);
    output.push(finalName);
  }
  return output;
}

/**
 * v3 → v4 wiring (one-way; v3 may call into the v4 barrel, v4 knows nothing of
 * v3). Resolves the scaffold template through the v4 distribution channel, then
 * renders the located template's entries onto disk.
 *
 * Expected failures from the v4 operations surface as thrown `FxError`s,
 * matching how the legacy local-template action rejects.
 *
 * `telemetryProps` (the `GenerateTemplate` event's props) is populated with the
 * resolved `source` as soon as resolution succeeds — before the package is read
 * or rendered — so a later digest/render failure still carries the origin and
 * version, making v4 errors attributable in telemetry.
 */
export async function scaffoldFromV4Channel(
  context: GeneratorContext,
  locator: TemplateLocator,
  telemetryProps?: Record<string, string>
): Promise<TemplateSource> {
  const channelConfig = {
    templatesV4TagListURL: templateConfig.templatesV4TagListURL,
    templateDownloadBaseURL: templateConfig.templateDownloadBaseURL,
    tryLimits: context.tryLimits ?? defaultTryLimits,
  };
  const port = v4TemplateBridgeDeps.createTemplateSourcePort(
    channelConfig,
    v4TemplateBridgeDeps.loadBundledFloor()
  );

  const sourceResult = await v4TemplateBridgeDeps.resolveTemplateSource({
    range: templateConfig.v4.range,
    bundled: templateConfig.v4.bundled,
    port,
  });
  if (sourceResult.isErr()) {
    throw sourceResult.error;
  }
  const source = sourceResult.value;
  merge(telemetryProps, {
    [TelemetryProperty.TemplatePackageSource]: source.origin,
    [TelemetryProperty.TemplatePackageVersion]: source.version,
    [TelemetryProperty.TemplatePackageDigest]: source.digest,
  });

  const bytesResult = v4TemplateBridgeDeps.loadResolvedPackage(source, port);
  if (bytesResult.isErr()) {
    throw bytesResult.error;
  }

  const entriesResult = v4TemplateBridgeDeps.openTemplatePackage(bytesResult.value, locator);
  if (entriesResult.isErr()) {
    throw entriesResult.error;
  }

  context.outputs = await renderTemplateEntries(context, entriesResult.value);
  return source;
}
