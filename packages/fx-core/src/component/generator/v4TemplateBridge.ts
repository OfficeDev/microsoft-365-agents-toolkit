// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import { merge } from "lodash";
import path from "path";
import templateConfig from "../../common/templates-config.json";
import { TelemetryProperty } from "../../common/telemetry";
import {
  TemplateFileEntry,
  TemplateLocator,
  TemplateSource,
  createTemplateSourcePort,
  loadBundledFloor,
  loadResolvedPackage,
  openTemplatePackage,
  resolveTemplateSource,
} from "../../v4";
import { defaultTryLimits } from "./constant";
import { GeneratorContext } from "./generatorAction";

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
    const filePath = path.join(context.destination, finalName);
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
  const port = createTemplateSourcePort(channelConfig, loadBundledFloor());

  const sourceResult = await resolveTemplateSource({
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

  const bytesResult = loadResolvedPackage(source, port);
  if (bytesResult.isErr()) {
    throw bytesResult.error;
  }

  const entriesResult = openTemplatePackage(bytesResult.value, locator);
  if (entriesResult.isErr()) {
    throw entriesResult.error;
  }

  context.outputs = await renderTemplateEntries(context, entriesResult.value);
  return source;
}
