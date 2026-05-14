// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { createDriver } from "../../createDriver";
import { buildAppPackage } from "../../../teamsApp/packageBuilder";

const inputSchema = z.object({
  /** Project root directory */
  projectPath: z.string().min(1),
  /** Optional explicit manifest path (auto-detected if omitted) */
  manifestPath: z.string().optional(),
  /** Output path for the .zip file */
  outputZipPath: z.string().min(1),
  /** Optional folder for resolved manifest files */
  outputFolder: z.string().optional(),
});

/**
 * Driver: teamsApp/zipAppPackage
 *
 * Builds a Teams app package (.zip) with full env-var resolution,
 * DA/plugin processing, and icon handling.
 *
 * Outputs:
 * - TEAMS_APP_PACKAGE_PATH: absolute path to the produced ZIP
 */
export const zipAppPackageDriver = createDriver({
  id: "teamsApp/zipAppPackage",
  name: "Zip App Package",
  inputSchema,
  execute: async (ctx, config) => {
    const result = await buildAppPackage(ctx, {
      projectPath: config.projectPath,
      manifestPath: config.manifestPath,
      outputZipPath: config.outputZipPath,
      outputFolder: config.outputFolder,
    });

    return result.map((r) => ({
      outputs: { TEAMS_APP_PACKAGE_PATH: r.zipPath },
    }));
  },
});
