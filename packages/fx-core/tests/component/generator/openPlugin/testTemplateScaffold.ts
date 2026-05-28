// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import * as path from "path";
import {
  renderTemplateFileData,
  renderTemplateFileName,
} from "../../../../src/component/generator/utils";

// Resolves to <repoRoot>/templates/vsc/common/open-plugin-import. Tests use this
// to scaffold the baseline directly from the source template tree instead of
// the (non-existent in dev) common.zip fallback.
const SOURCE_TEMPLATE_DIR = path.resolve(
  __dirname,
  "..",
  "..",
  "..",
  "..",
  "..",
  "..",
  "templates",
  "vsc",
  "common",
  "open-plugin-import"
);

export async function scaffoldOpenPluginTemplateFromSource(
  destinationPath: string,
  replaceMap: Record<string, string>
): Promise<void> {
  if (!(await fs.pathExists(SOURCE_TEMPLATE_DIR))) {
    throw new Error(`open-plugin-import template source not found at ${SOURCE_TEMPLATE_DIR}`);
  }
  await copyTree(SOURCE_TEMPLATE_DIR, destinationPath, replaceMap);
}

async function copyTree(
  srcDir: string,
  dstDir: string,
  replaceMap: Record<string, string>
): Promise<void> {
  const entries = await fs.readdir(srcDir, { withFileTypes: true });
  for (const entry of entries) {
    const srcPath = path.join(srcDir, entry.name);
    if (entry.isDirectory()) {
      const dstSubDir = path.join(dstDir, entry.name);
      await fs.ensureDir(dstSubDir);
      await copyTree(srcPath, dstSubDir, replaceMap);
      continue;
    }
    const data = await fs.readFile(srcPath);
    const renderedName = renderTemplateFileName(entry.name, data, replaceMap);
    const renderedData = renderTemplateFileData(entry.name, data, replaceMap);
    const dstPath = path.join(dstDir, renderedName);
    await fs.ensureDir(path.dirname(dstPath));
    await fs.writeFile(dstPath, renderedData);
  }
}
