// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError, UserError } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import { createContext } from "../../../common/globalVars";
import { Generator } from "../generator";
import { TemplateNames } from "../templates/templateNames";
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
 *
 * Static baseline files (m365agents.yml, README, .gitignore, .vscode, env)
 * come from the `open-plugin-convert` template, which ships in the standard
 * template release pipeline and can be updated independently of fx-core.
 * Variable-length outputs (manifest, skill folders, icons) are written as
 * the post-scaffold step here.
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

    // 1. Scaffold the static baseline from the open-plugin-convert template.
    const ctx = createContext();
    ctx.templateVariables = { appName: parsed.manifest.name };
    const scaffoldRes = await Generator.generateTemplate(
      ctx,
      projectPath,
      TemplateNames.OpenPluginConvert,
      "common"
    );
    if (scaffoldRes.isErr()) return err(scaffoldRes.error);

    // 2. Post-scaffold: write the dynamic outputs.
    const appPackageDir = path.join(projectPath, "appPackage");
    await fs.ensureDir(appPackageDir);

    // Manifest (vDevPreview agentSkills/agentConnectors are variable-length).
    await fs.writeJSON(path.join(appPackageDir, "manifest.json"), manifest, { spaces: 4 });

    // Copy skill folders and (when present) the commands folder.
    for (const op of copyOps) {
      await fs.copy(op.src, path.join(projectPath, op.destRelative));
    }

    // Icons.
    await applyIcons(parsed, appPackageDir, warnings);

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
