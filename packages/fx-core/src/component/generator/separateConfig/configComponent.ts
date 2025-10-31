// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import mustache from "mustache";
import * as os from "os";
import * as path from "path";
import { getTemplatesFolder } from "../../../folder";
import { getTemplateReplaceMap } from "../templates/templateReplaceMap";
import { TemplateConfig } from "./interface";
import { mergeJsonFile } from "./jsonMerge";
import { mergeYamlFile } from "./yamlMerge";

export async function scaffoldConfigComponents(
  templateConfig: TemplateConfig,
  destinationPath: string,
  inputs: Inputs
): Promise<void> {
  const templatesConfigsRoot = path.join(getTemplatesFolder(), "configs");
  if (!(await fs.pathExists(templatesConfigsRoot))) {
    return; // nothing to do
  }

  // Stage 1: collect & merge into temp folder
  const tempFolder = await fs.mkdtemp(path.join(os.tmpdir(), "teamsfx-config-"));
  const features = templateConfig.features || {};
  for (const configComponent of templateConfig.components) {
    features[configComponent.name] = true;
  }
  for (const configComponent of templateConfig.components) {
    const configPath = path.join(templatesConfigsRoot, configComponent.name);
    await mergeConfigFiles(configPath, tempFolder, features, inputs);
  }

  // Stage 2: copy from temp folder to final destination
  fs.rmdirSync(path.join(tempFolder, "rendered"), { recursive: true }); // clean up rendered files
  await renderAndCopy(tempFolder, destinationPath);
}

async function mergeConfigFiles(
  configPath: string,
  outputDir: string,
  features: Record<string, unknown>,
  inputs: Inputs
): Promise<void> {
  const files = await getAllFiles(configPath);
  const replaceMap = getTemplateReplaceMap(inputs);

  for (const file of files) {
    let relativePath = path.relative(configPath, file);
    let configSourceFile = file;
    let targetPath = path.join(outputDir, relativePath);

    await fs.ensureDir(path.dirname(targetPath));

    if (configSourceFile.endsWith(".mu")) {
      // use mustache to render feature flags
      const renderedFile = path.join(outputDir, "rendered", path.basename(configSourceFile, ".mu"));
      await renderMustacheTemplate(configSourceFile, renderedFile, features);
      relativePath = relativePath.slice(0, -3); // remove .mu extension
      targetPath = targetPath.slice(0, -3); // remove .mu extension
      configSourceFile = renderedFile;
    }

    if (configSourceFile.endsWith(".tpl")) {
      const tempDir = path.join(outputDir, "rendered");
      await fs.ensureDir(tempDir);
      const renderedFile = path.join(tempDir, path.basename(configSourceFile, ".tpl")); // remove .tpl
      await renderMustacheTemplateWithMap(configSourceFile, renderedFile, replaceMap);
      targetPath = targetPath.slice(0, -4); // remove .tpl suffix
      configSourceFile = renderedFile;
    }

    if (relativePath.endsWith(".json.tpl") || relativePath.endsWith(".json")) {
      await mergeJsonFile(configSourceFile, targetPath);
    } else if (relativePath.endsWith(".yml.tpl") || relativePath.endsWith(".yml")) {
      await mergeYamlFile(configSourceFile, targetPath);
    } else {
      await appendFile(configSourceFile, targetPath);
    }
  }
}

async function getAllFiles(dir: string): Promise<string[]> {
  const files: string[] = [];
  const items = await fs.readdir(dir);

  for (const item of items) {
    const fullPath = path.join(dir, item);
    const stat = await fs.stat(fullPath);

    if (stat.isDirectory()) {
      files.push(...(await getAllFiles(fullPath)));
    } else {
      files.push(fullPath);
    }
  }

  return files;
}

async function renderMustacheTemplate(
  file: string,
  renderedFile: string,
  features: Record<string, any>
) {
  const templateContent = await fs.readFile(file, "utf8");
  mustache.tags = ["<<", ">>"];

  const rendered = mustache.render(templateContent, features);

  await fs.ensureDir(path.dirname(renderedFile));
  await fs.writeFile(renderedFile, rendered, "utf8");
}

async function appendFile(sourcePath: string, targetPath: string): Promise<void> {
  const sourceContent = await fs.readFile(sourcePath, "utf-8");

  if (await fs.pathExists(targetPath)) {
    // Append to existing file
    const targetContent = await fs.readFile(targetPath, "utf-8");
    const mergedContent = targetContent + "\n" + sourceContent;
    await fs.writeFile(targetPath, mergedContent);
  } else {
    // Copy as new file
    await fs.writeFile(targetPath, sourceContent);
  }
}

async function renderAndCopy(sourceFolder: string, destinationFolder: string): Promise<void> {
  const files = await getAllFiles(sourceFolder);

  for (const file of files) {
    const relativePath = path.relative(sourceFolder, file);
    const fileName = path.basename(file);

    // Construct destination path with processed filename
    const relativeDir = path.dirname(relativePath);
    const destinationPath = path.join(destinationFolder, relativeDir, fileName);

    await fs.ensureDir(path.dirname(destinationPath));

    // Step 2: Merge if file exists, otherwise copy
    if (await fs.pathExists(destinationPath)) {
      // Merge with existing file
      if (fileName.endsWith(".json")) {
        await mergeJsonFile(file, destinationPath);
      } else if (fileName.endsWith(".yml")) {
        await mergeYamlFile(file, destinationPath);
      } else {
        // For other files, append content
        await appendFile(file, destinationPath);
      }
    } else {
      // Copy file if it doesn't exist
      await fs.copy(file, destinationPath);
    }
  }
}

async function renderMustacheTemplateWithMap(
  file: string,
  renderedFile: string,
  replaceMap: { [key: string]: string }
) {
  let templateContent = await fs.readFile(file, "utf8");

  // Temporarily replace ${{xxx}} patterns with a placeholder to protect them from mustache
  const placeholder = "___DOLLAR_MUSTACHE___";
  const dollarPatterns: Array<{ index: number; value: string }> = [];
  templateContent = templateContent.replace(/\$\{\{([^}]+)\}\}/g, (match) => {
    dollarPatterns.push({ index: dollarPatterns.length, value: match });
    return `${placeholder}${dollarPatterns.length - 1}${placeholder}`;
  });

  // Render with standard mustache delimiters
  mustache.tags = ["{{", "}}"];
  let rendered = mustache.render(templateContent, replaceMap);

  // Restore the ${{xxx}} patterns
  rendered = rendered.replace(
    new RegExp(`${placeholder}(\\d+)${placeholder}`, "g"),
    (_, index) => dollarPatterns[parseInt(index)].value
  );

  await fs.ensureDir(path.dirname(renderedFile));
  await fs.writeFile(renderedFile, rendered, "utf8");
}
