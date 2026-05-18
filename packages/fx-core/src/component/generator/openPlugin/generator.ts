// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError, UserError } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import {
  ENV_DEV_CONTENT,
  GITIGNORE_CONTENT,
  renderM365AgentsYml,
  renderReadme,
  VSCODE_EXTENSIONS_JSON,
  VSCODE_LAUNCH_JSON,
  VSCODE_SETTINGS_JSON,
} from "./assets";
import { applyIcons } from "./iconStrategy";
import { mapToTtkProject } from "./mapper";
import { readOpenPluginDir } from "./parser";
import { ConvertInputs } from "./types";

export const OPEN_PLUGIN_CONVERT_SOURCE = "OpenPluginConvert";

export interface ConvertResult {
  projectPath: string;
  warnings: string[];
}

/**
 * Convert an Open Plugin / Claude Code plugin / Cursor plugin directory into
 * a scaffolded Teams Toolkit project. The output is a usable TTK project; run
 * `atk teamsapp package` from inside it to produce the upload zip.
 */
export async function convertOpenPlugin(
  inputs: ConvertInputs
): Promise<Result<ConvertResult, FxError>> {
  try {
    if (!inputs.path) {
      return err(
        new UserError(OPEN_PLUGIN_CONVERT_SOURCE, "MissingPluginPath", "--path is required.")
      );
    }
    const parsed = await readOpenPluginDir(inputs.path);
    const { manifest, copyOps, warnings } = mapToTtkProject(parsed, inputs);

    const defaultOutput = path.join(process.cwd(), parsed.manifest.name);
    const projectPath = path.resolve(inputs.output ?? defaultOutput);

    if (await fs.pathExists(projectPath)) {
      const entries = await fs.readdir(projectPath);
      if (entries.length > 0) {
        return err(
          new UserError(
            OPEN_PLUGIN_CONVERT_SOURCE,
            "OutputDirectoryNotEmpty",
            `Output directory is not empty: ${projectPath}. Choose a different --output path or empty the directory.`
          )
        );
      }
    }

    const appPackageDir = path.join(projectPath, "appPackage");
    await fs.ensureDir(appPackageDir);

    // 1. Manifest (programmatic; vDevPreview agentSkills/agentConnectors are variable-length).
    await fs.writeJSON(path.join(appPackageDir, "manifest.json"), manifest, { spaces: 4 });

    // 2. Copy skill folders and (when present) the commands folder.
    for (const op of copyOps) {
      await fs.copy(op.src, path.join(projectPath, op.destRelative));
    }

    // 3. Icons.
    await applyIcons(parsed, appPackageDir, warnings);

    // 4. Static baseline files.
    await writeStaticBaseline(projectPath, parsed.manifest.name);

    return ok({ projectPath, warnings });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return err(
      new SystemError({
        source: OPEN_PLUGIN_CONVERT_SOURCE,
        name: "ConvertOpenPluginFailed",
        message,
        displayMessage: message,
      })
    );
  }
}

async function writeStaticBaseline(projectPath: string, appName: string): Promise<void> {
  await fs.writeFile(path.join(projectPath, ".gitignore"), GITIGNORE_CONTENT, "utf8");
  await fs.writeFile(path.join(projectPath, "README.md"), renderReadme(appName), "utf8");
  await fs.writeFile(
    path.join(projectPath, "m365agents.yml"),
    renderM365AgentsYml(appName),
    "utf8"
  );

  const envDir = path.join(projectPath, "env");
  await fs.ensureDir(envDir);
  await fs.writeFile(path.join(envDir, ".env.dev"), ENV_DEV_CONTENT, "utf8");

  const vscodeDir = path.join(projectPath, ".vscode");
  await fs.ensureDir(vscodeDir);
  await fs.writeFile(path.join(vscodeDir, "launch.json"), VSCODE_LAUNCH_JSON, "utf8");
  await fs.writeFile(path.join(vscodeDir, "settings.json"), VSCODE_SETTINGS_JSON, "utf8");
  await fs.writeFile(path.join(vscodeDir, "extensions.json"), VSCODE_EXTENSIONS_JSON, "utf8");
}
