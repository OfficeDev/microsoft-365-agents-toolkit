// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { execFile } from "node:child_process";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { AzureArmClient } from "../../../clients/azure/client";
import { azureManagementScopes, ArmDeploymentRequest } from "../../../clients/azure/types";
import { resolveEnvPlaceholders, getEnvVariables } from "../../../manifest/resolve";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const SOURCE = "arm/deploy";

const templateSchema = z.object({
  path: z.string().min(1),
  parameters: z.string().optional(),
  deploymentName: z.string().min(1),
});

const inputSchema = z.object({
  subscriptionId: z.string().regex(UUID_RE, "subscriptionId must be a valid UUID"),
  resourceGroupName: z.string().min(1),
  templates: z.array(templateSchema).min(1),
  bicepCliVersion: z.string().optional(),
});

type ArmDeployConfig = z.infer<typeof inputSchema>;

/**
 * Driver: arm/deploy
 *
 * Deploys one or more ARM/Bicep templates to an Azure resource group.
 * - Compiles .bicep files to JSON via the Bicep CLI
 * - Deploys all templates in parallel
 * - Extracts deployment outputs as uppercase env vars (nested keys joined with __)
 */
export const armDeployDriver = createDriver<ArmDeployConfig>({
  id: "arm/deploy",
  name: "Deploy ARM Templates",
  inputSchema,
  execute: async (ctx, config) => {
    if (!ctx.projectPath) {
      return err(
        userError("MissingProjectPath", "projectPath is required for ARM deployment", {
          source: SOURCE,
        })
      );
    }

    // Acquire Azure token
    const credential = await ctx.auth.azureAccountProvider.getIdentityCredentialAsync();
    if (!credential) {
      return err(
        systemError("AzureCredentialError", "Failed to acquire Azure credential", {
          source: SOURCE,
        })
      );
    }
    const tokenResult = await credential.getToken(azureManagementScopes());
    if (!tokenResult) {
      return err(
        systemError("AzureTokenError", "Failed to acquire Azure management token", {
          source: SOURCE,
        })
      );
    }

    const client = new AzureArmClient(ctx, tokenResult.token);
    const bicepCommand = config.bicepCliVersion ? undefined : "bicep";

    // Deploy all templates in parallel
    const deploymentPromises = config.templates.map(async (template) => {
      ctx.logger.info(`[${SOURCE}] Deploying '${template.deploymentName}' from ${template.path}`);

      // Load and compile template
      const templateJson = await loadTemplate(
        template.path,
        ctx.projectPath!,
        bicepCommand ?? "bicep"
      );

      // Load parameters if specified
      let parameters: Record<string, unknown> | null = null;
      if (template.parameters) {
        parameters = await loadParameters(template.parameters, ctx.projectPath!);
      }

      const body: ArmDeploymentRequest = {
        properties: {
          template: templateJson,
          parameters: parameters,
          mode: "Incremental",
        },
      };

      const res = await client.deployTemplate(
        config.subscriptionId,
        config.resourceGroupName,
        template.deploymentName,
        body
      );

      if (res.isErr()) return res;
      ctx.logger.info(`[${SOURCE}] Deployment '${template.deploymentName}' succeeded`);
      return res;
    });

    const results = await Promise.all(deploymentPromises);

    // Check for failures
    for (const result of results) {
      if (result.isErr()) return err(result.error);
    }

    // Extract and merge outputs from all deployments
    const outputs = new Map<string, string>();
    for (const result of results) {
      if (result.isOk() && result.value.properties?.outputs) {
        convertOutputs(result.value.properties.outputs, outputs);
      }
    }

    return ok({ outputs: Object.fromEntries(outputs) });
  },
});

/**
 * Load a template file. If .bicep, compile to JSON via CLI.
 */
async function loadTemplate(
  templatePath: string,
  projectPath: string,
  bicepCommand: string
): Promise<Record<string, unknown>> {
  const absPath = path.isAbsolute(templatePath)
    ? templatePath
    : path.join(projectPath, templatePath);

  const ext = path.extname(absPath).toLowerCase();
  if (ext === ".bicep") {
    return compileBicep(absPath, bicepCommand);
  }

  const content = await fs.readFile(absPath, "utf-8");
  return JSON.parse(content);
}

/**
 * Compile a Bicep file to JSON ARM template using the Bicep CLI.
 */
function compileBicep(filePath: string, command: string): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    execFile(command, ["build", filePath, "--stdout"], { timeout: 120_000 }, (error, stdout) => {
      if (error) {
        reject(
          userError(
            "CompileBicepError",
            `Failed to compile Bicep '${filePath}': ${error.message}`,
            {
              source: SOURCE,
              inner: error,
            }
          )
        );
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (_parseErr) {
        reject(
          userError("CompileBicepError", `Bicep output for '${filePath}' is not valid JSON`, {
            source: SOURCE,
          })
        );
      }
    });
  });
}

/**
 * Load and resolve a parameters JSON file.
 * Expands ${{VAR}} placeholders from environment variables.
 */
async function loadParameters(
  parametersPath: string,
  projectPath: string
): Promise<Record<string, unknown>> {
  const absPath = path.isAbsolute(parametersPath)
    ? parametersPath
    : path.join(projectPath, parametersPath);

  const content = await fs.readFile(absPath, "utf-8");
  const { content: resolved } = resolveEnvPlaceholders(content);

  // Check for unresolved placeholders
  const unresolvedVars = getEnvVariables(resolved);
  if (unresolvedVars.length > 0) {
    throw userError(
      "MissingEnvironmentVariables",
      `Unresolved environment variables in '${parametersPath}': ${unresolvedVars.join(", ")}`,
      { source: SOURCE }
    );
  }

  const parsed = JSON.parse(resolved);
  return parsed.parameters ?? parsed;
}

/**
 * Convert ARM deployment outputs to a flat string map.
 * Keys are uppercased; nested objects use __ as separator.
 */
function convertOutputs(
  outputs: Record<string, { type: string; value: unknown }>,
  map: Map<string, string>,
  prefix?: string
): void {
  for (const [key, entry] of Object.entries(outputs)) {
    const mapKey = prefix ? `${prefix}__${key}` : key;
    if (entry.value && typeof entry.value === "object" && !Array.isArray(entry.value)) {
      convertOutputs(entry.value as Record<string, { type: string; value: unknown }>, map, mapKey);
    } else {
      const upperKey = mapKey.toUpperCase();
      map.set(upperKey, entry.value != null ? String(entry.value) : "");
    }
  }
}
