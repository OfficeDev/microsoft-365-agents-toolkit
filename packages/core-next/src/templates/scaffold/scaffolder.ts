// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import { Result, ok, err } from "neverthrow";
import type { AtkContext } from "../../core/context";
import type { AtkError } from "../../core/error";
import { userError, systemError } from "../../core/error";
import type { ScaffoldContext, TemplateInfo, TemplateConfig } from "./types";
import { renderTemplateFileName, renderTemplateFileData } from "./render";
import { resolveTemplateUrl, fetchZip, unzipWithTransform, loadLocalFallback } from "./download";
import { getTemplatesFolder } from "../../folder";

/**
 * Default template configuration. Can be overridden via environment variables.
 */
const DEFAULT_TEMPLATE_CONFIG: TemplateConfig = {
  downloadBaseUrl: "https://github.com/OfficeDev/Microsoft-365-Agents-Toolkit/releases/download",
  tagPrefix: "templates@",
  archiveExt: ".zip",
  version: process.env.TEMPLATE_VERSION ?? "local",
};

/**
 * Resolve the fallback directory for local template zip files.
 * Priority: explicit parameter > TEMPLATE_FALLBACK_DIR env var > bundled templates/fallback.
 */
function resolveFallbackDir(explicit?: string): string | undefined {
  return (
    explicit ?? process.env.TEMPLATE_FALLBACK_DIR ?? path.join(getTemplatesFolder(), "fallback")
  );
}

/**
 * Scaffold one or more templates into a destination directory.
 *
 * For each TemplateInfo:
 * 1. Attempt remote download (unless version is "local")
 * 2. Fall back to local zip if remote fails or is disabled
 * 3. Unzip with Mustache rendering of .tpl files and file names
 */
export async function scaffoldTemplates(
  ctx: AtkContext,
  templates: TemplateInfo[],
  destinationPath: string,
  config?: Partial<TemplateConfig>,
  fallbackDir?: string
): Promise<Result<string[], AtkError>> {
  const templateConfig = { ...DEFAULT_TEMPLATE_CONFIG, ...config };
  const effectiveFallbackDir = resolveFallbackDir(fallbackDir);
  const allOutputs: string[] = [];

  for (const tplInfo of templates) {
    const subDest = tplInfo.subFolder ? `${destinationPath}/${tplInfo.subFolder}` : destinationPath;

    const replaceMap = tplInfo.replaceMap ?? {};
    const tplName = tplInfo.templateName;
    const tplPrefix = `${tplName}/`;

    // Compose filter: only entries under the template folder, then user filter
    const composedFilter = (name: string): boolean => {
      const normalised = name.replace(/\\/g, "/");
      if (!normalised.startsWith(tplPrefix)) return false;
      return !tplInfo.filterFn || tplInfo.filterFn(normalised);
    };

    // Compose name transform: strip template folder prefix, then render Mustache
    const composedNameFn = (name: string, data: Buffer): string => {
      const normalised = name.replace(/\\/g, "/");
      const stripped = normalised.startsWith(tplPrefix)
        ? normalised.slice(tplPrefix.length)
        : normalised;
      return renderTemplateFileName(stripped, data, replaceMap);
    };

    const scaffoldCtx: ScaffoldContext = {
      name: tplName,
      language: tplInfo.language,
      destination: subDest,
      tryLimits: 1,
      timeoutInMs: 20000,
      useFallback: false,
      outputs: [],
      fileNameReplaceFn: composedNameFn,
      fileDataReplaceFn: (name, data) => renderTemplateFileData(name, data, replaceMap),
      filterFn: composedFilter,
    };

    // Step 1: Try remote download
    const remoteUrl = resolveTemplateUrl(templateConfig, tplInfo.language);
    let zipBuffer: Buffer | undefined;

    if (remoteUrl) {
      ctx.logger.debug(
        `[scaffold] Downloading template "${tplInfo.templateName}" from ${remoteUrl}`
      );
      zipBuffer = await fetchZip(remoteUrl, scaffoldCtx.tryLimits, scaffoldCtx.timeoutInMs);
    }

    // Step 2: Fall back to local
    if (!zipBuffer && effectiveFallbackDir) {
      ctx.logger.debug(
        `[scaffold] Using local fallback for "${tplInfo.templateName}" from ${effectiveFallbackDir}`
      );
      zipBuffer = await loadLocalFallback(effectiveFallbackDir, tplInfo.language);
      scaffoldCtx.useFallback = true;
    }

    if (!zipBuffer) {
      return err(
        userError(
          "TemplateNotFound",
          `Template "${tplInfo.templateName}" (${tplInfo.language}) could not be downloaded or found locally.`,
          { source: "templates/scaffold" }
        )
      );
    }

    // Step 3: Unzip with Mustache transforms
    try {
      const files = await unzipWithTransform(zipBuffer, scaffoldCtx);
      allOutputs.push(...files);
      ctx.logger.debug(`[scaffold] Scaffolded ${files.length} files for "${tplInfo.templateName}"`);
    } catch (e) {
      return err(
        systemError("ScaffoldFailed", `Failed to scaffold "${tplInfo.templateName}": ${e}`, {
          source: "templates/scaffold",
          inner: e instanceof Error ? e : new Error(String(e)),
        })
      );
    }
  }

  return ok(allOutputs);
}
