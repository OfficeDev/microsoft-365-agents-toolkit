// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import * as path from "path";
import { generatePlaceholderPng } from "./placeholderPng";
import { ParsedOpenPlugin } from "./types";

const COLOR_FILL: [number, number, number] = [0x4a, 0x90, 0xd9];
const OUTLINE_FILL: [number, number, number] = [0xff, 0xff, 0xff];

/**
 * Write color.png and outline.png into the appPackage directory. Resolution
 * order matches the conversion plan: user-supplied root PNG → Open Plugin
 * `logo` (when it is a local .png) → generated placeholder.
 */
export async function applyIcons(
  parsed: ParsedOpenPlugin,
  appPackageDir: string,
  warnings: string[]
): Promise<void> {
  const colorDest = path.join(appPackageDir, "color.png");
  const outlineDest = path.join(appPackageDir, "outline.png");

  if (parsed.hasColorPng) {
    await fs.copy(path.join(parsed.pluginRoot, "color.png"), colorDest);
  } else if (typeof parsed.manifest.logo === "string" && parsed.manifest.logo) {
    const applied = await tryApplyLogo(parsed, colorDest, warnings);
    if (!applied) {
      await fs.writeFile(colorDest, generatePlaceholderPng(192, ...COLOR_FILL));
    }
  } else {
    await fs.writeFile(colorDest, generatePlaceholderPng(192, ...COLOR_FILL));
  }

  if (parsed.hasOutlinePng) {
    await fs.copy(path.join(parsed.pluginRoot, "outline.png"), outlineDest);
  } else {
    await fs.writeFile(outlineDest, generatePlaceholderPng(32, ...OUTLINE_FILL));
  }
}

async function tryApplyLogo(
  parsed: ParsedOpenPlugin,
  colorDest: string,
  warnings: string[]
): Promise<boolean> {
  const logo = parsed.manifest.logo as string;
  if (/^https?:\/\//i.test(logo)) {
    warnings.push(
      `'logo' field points to a remote URL (${logo}); using placeholder color.png. Download manually if you want to ship the original.`
    );
    return false;
  }
  if (!/\.png$/i.test(logo)) {
    warnings.push(`'logo' field '${logo}' is not a .png file; using placeholder color.png.`);
    return false;
  }
  const logoAbs = path.resolve(parsed.pluginRoot, logo);
  const relativeToRoot = path.relative(parsed.pluginRoot, logoAbs);
  if (relativeToRoot.startsWith("..") || path.isAbsolute(relativeToRoot)) {
    warnings.push(
      `'logo' field '${logo}' resolves outside the plugin root; using placeholder color.png.`
    );
    return false;
  }
  if (!(await fs.pathExists(logoAbs))) {
    warnings.push(`'logo' field '${logo}' does not exist; using placeholder color.png.`);
    return false;
  }
  await fs.copy(logoAbs, colorDest);
  return true;
}
