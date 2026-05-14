// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { createDriver } from "../../createDriver";
import { zipDeployExecute } from "../azureAppService/zipDeploy";

const inputSchema = z.object({
  resourceId: z.string().min(1),
  artifactFolder: z.string().min(1),
  workingDirectory: z.string().optional(),
  ignoreFile: z.string().optional(),
  dryRun: z.boolean().optional(),
  outputZipFile: z.string().optional(),
});

/**
 * Driver: azureFunctions/zipDeploy
 *
 * Identical to azureAppService/zipDeploy but with Functions-specific naming.
 * Both use the same SCM zip deploy API under the hood.
 */
export const azureFunctionsZipDeployDriver = createDriver({
  id: "azureFunctions/zipDeploy",
  name: "Deploy to Azure Functions",
  inputSchema,
  execute: async (ctx, config) => {
    return zipDeployExecute(ctx, config, "azureFunctions/zipDeploy", true);
  },
});
