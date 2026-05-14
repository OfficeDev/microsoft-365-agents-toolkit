// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err, Result } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import AdmZip from "adm-zip";
import axios from "axios";
import { createDriver } from "../../createDriver";
import { AtkError, systemError, userError } from "../../../core/error";
import { AtkContext } from "../../../core/context";
import { AzureArmClient } from "../../../clients/azure/client";
import { azureManagementScopes, AzureResourceId, DeployStatus } from "../../../clients/azure/types";
import { DriverOutput } from "../../types";

const SOURCE = "azureAppService/zipDeploy";
const DEPLOY_CHECK_INTERVAL_MS = 10_000; // 10 seconds
const DEPLOY_CHECK_MAX_ATTEMPTS = 120; // ~20 minutes
const DEPLOY_UPLOAD_RETRIES = 2;

const RESOURCE_ID_PATTERN =
  /\/subscriptions\/([^/]*)\/resourceGroups\/([^/]*)\/providers\/Microsoft\.Web\/sites\/([^/]*)/i;

const inputSchema = z.object({
  /** Azure resource ID for the App Service */
  resourceId: z.string().min(1),
  /** Artifact folder path (relative to projectPath) */
  artifactFolder: z.string().min(1),
  /** Working directory (defaults to ".") */
  workingDirectory: z.string().optional(),
  /** Ignore file path (.gitignore-style, relative to workingDirectory) */
  ignoreFile: z.string().optional(),
  /** Dry-run: create zip only, don't deploy */
  dryRun: z.boolean().optional(),
  /** Custom zip output path */
  outputZipFile: z.string().optional(),
});

export type ZipDeployConfig = z.infer<typeof inputSchema>;

/**
 * Driver: azureAppService/zipDeploy
 *
 * Creates a ZIP of the artifact folder and deploys it to Azure App Service
 * via the SCM zip deploy API.
 */
export const azureAppServiceZipDeployDriver = createDriver<ZipDeployConfig>({
  id: "azureAppService/zipDeploy",
  name: "Deploy to Azure App Service",
  inputSchema,
  execute: async (ctx, config) => {
    return zipDeployExecute(ctx, config, SOURCE, true);
  },
});

/**
 * Shared zip deploy logic used by both App Service and Functions drivers.
 */
export async function zipDeployExecute(
  ctx: AtkContext,
  config: ZipDeployConfig,
  source: string,
  restartAfterDeploy: boolean
): Promise<Result<DriverOutput, AtkError>> {
  if (!ctx.projectPath) {
    return err(userError("MissingProjectPath", "projectPath is required for deploy", { source }));
  }

  // Parse resource ID
  const resourceId = parseResourceId(config.resourceId, source);
  if (resourceId.isErr()) return err(resourceId.error);
  const { subscriptionId, resourceGroupName, instanceId } = resourceId.value;

  // Resolve paths
  const workDir = path.isAbsolute(config.workingDirectory ?? ".")
    ? config.workingDirectory ?? "."
    : path.join(ctx.projectPath, config.workingDirectory ?? ".");
  const distDir = path.isAbsolute(config.artifactFolder)
    ? config.artifactFolder
    : path.join(workDir, config.artifactFolder);

  // Create ZIP (readdir inside createZip will fail with ENOENT if distDir is missing)
  ctx.logger.info(`[${source}] Creating ZIP from ${distDir}`);
  const zipPath = config.outputZipFile
    ? path.isAbsolute(config.outputZipFile)
      ? config.outputZipFile
      : path.join(workDir, config.outputZipFile)
    : path.join(workDir, ".deployment", "deployment.zip");

  const zipDir = path.dirname(zipPath);
  await fs.mkdir(zipDir, { recursive: true });

  const zipResult = await createZip(distDir, zipPath, config.ignoreFile, workDir);
  if (zipResult.isErr()) return err(zipResult.error);

  if (config.dryRun) {
    ctx.logger.info(`[${source}] Dry run complete. ZIP at: ${zipPath}`);
    return ok({ outputs: {} });
  }

  // Acquire Azure token
  const credential = await ctx.auth.azureAccountProvider.getIdentityCredentialAsync();
  if (!credential) {
    return err(
      systemError("AzureCredentialError", "Failed to acquire Azure credential", { source })
    );
  }
  const tokenResult = await credential.getToken(azureManagementScopes());
  if (!tokenResult) {
    return err(
      systemError("AzureTokenError", "Failed to acquire Azure management token", { source })
    );
  }

  const armClient = new AzureArmClient(ctx, tokenResult.token);

  // Get SCM endpoint
  const scmResult = await armClient.getScmEndpoint(subscriptionId, resourceGroupName, instanceId);
  if (scmResult.isErr()) return err(scmResult.error);
  const scmEndpoint = `${scmResult.value}/api/zipdeploy?isAsync=true`;

  // Upload ZIP
  ctx.logger.info(`[${source}] Uploading ZIP to ${scmEndpoint}`);
  const zipBuffer = await fs.readFile(zipPath);

  // Validate ZIP magic bytes before uploading
  const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
  if (zipBuffer.length < 4 || zipBuffer.subarray(0, 4).compare(ZIP_MAGIC) !== 0) {
    return err(
      systemError("InvalidZipPackage", `ZIP file at '${zipPath}' is not a valid ZIP archive`, {
        source,
      })
    );
  }

  const uploadResult = await uploadZipWithRetry(
    scmEndpoint,
    zipBuffer,
    tokenResult.token,
    DEPLOY_UPLOAD_RETRIES,
    source
  );
  if (uploadResult.isErr()) return err(uploadResult.error);
  const locationUrl = uploadResult.value;

  // Poll deployment status
  ctx.logger.info(`[${source}] Polling deployment status...`);
  const statusResult = await checkDeployStatus(locationUrl, tokenResult.token, source);
  if (statusResult.isErr()) return err(statusResult.error);

  // Restart site
  if (restartAfterDeploy) {
    const restartRes = await armClient.restartSite(subscriptionId, resourceGroupName, instanceId);
    if (restartRes.isErr()) {
      ctx.logger.warning(`[${source}] Failed to restart site: ${restartRes.error.message}`);
    }
  }

  // Cleanup ZIP
  try {
    await fs.unlink(zipPath);
  } catch {
    // Non-fatal
  }

  ctx.logger.info(`[${source}] Deployment to ${instanceId} completed successfully`);
  return ok({ outputs: {} });
}

/**
 * Parse an Azure resource ID using the Web/sites pattern.
 */
function parseResourceId(resourceId: string, source: string): Result<AzureResourceId, AtkError> {
  const match = RESOURCE_ID_PATTERN.exec(resourceId);
  if (!match) {
    return err(
      userError("InvalidResourceId", `Resource ID does not match expected pattern: ${resourceId}`, {
        source,
      })
    );
  }
  return ok({
    subscriptionId: match[1],
    resourceGroupName: match[2],
    instanceId: match[3],
  });
}

/**
 * Create a ZIP file from the given directory.
 */
async function createZip(
  sourceDir: string,
  zipPath: string,
  ignoreFile: string | undefined,
  workDir: string
): Promise<Result<void, AtkError>> {
  try {
    // Load ignore patterns
    const ignorePatterns: string[] = [".deployment"];
    if (ignoreFile) {
      const ignoreFilePath = path.join(workDir, ignoreFile);
      try {
        const content = await fs.readFile(ignoreFilePath, "utf-8");
        content
          .split("\n")
          .map((line) => line.trim())
          .filter((line) => line && !line.startsWith("#"))
          .forEach((pattern) => ignorePatterns.push(pattern));
      } catch {
        // Ignore file not found — proceed without it
      }
    }

    const zip = new AdmZip();
    await addDirectoryToZip(zip, sourceDir, "", ignorePatterns);
    zip.writeZip(zipPath);
    return ok(undefined);
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") {
      return err(
        userError("ArtifactFolderNotFound", `Artifact folder '${sourceDir}' does not exist`, {
          source: "zipDeploy",
        })
      );
    }
    return err(
      systemError(
        "ZipCreationError",
        `Failed to create ZIP: ${e instanceof Error ? e.message : String(e)}`,
        {
          source: "zipDeploy",
        }
      )
    );
  }
}

/**
 * Recursively add directory contents to a ZIP, respecting ignore patterns.
 */
async function addDirectoryToZip(
  zip: AdmZip,
  baseDir: string,
  relativePath: string,
  ignorePatterns: string[]
): Promise<void> {
  const fullPath = relativePath ? path.join(baseDir, relativePath) : baseDir;
  const entries = await fs.readdir(fullPath, { withFileTypes: true });

  for (const entry of entries) {
    const entryRelative = relativePath ? path.join(relativePath, entry.name) : entry.name;

    // Simple pattern check — matches exact folder/file names in ignore list
    if (ignorePatterns.some((p) => entry.name === p || entryRelative === p)) {
      continue;
    }

    if (entry.isDirectory()) {
      await addDirectoryToZip(zip, baseDir, entryRelative, ignorePatterns);
    } else if (entry.isFile()) {
      const content = await fs.readFile(path.join(baseDir, entryRelative));
      zip.addFile(entryRelative.replace(/\\/g, "/"), content);
    }
  }
}

/**
 * Upload ZIP to SCM endpoint with retry on 5xx errors.
 */
async function uploadZipWithRetry(
  endpoint: string,
  zipBuffer: Buffer,
  token: string,
  maxRetries: number,
  source: string
): Promise<Result<string, AtkError>> {
  let lastError: unknown;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const res = await axios.post(endpoint, zipBuffer, {
        headers: {
          "Content-Type": "application/octet-stream",
          "Cache-Control": "no-cache",
          Authorization: `Bearer ${token}`,
        },
        maxContentLength: Infinity,
        maxBodyLength: Infinity,
        timeout: 600_000, // 10 minutes
        validateStatus: (status) => status < 500,
      });

      if (res.status === 200 || res.status === 202) {
        const location = res.headers["location"];
        if (!location) {
          return err(
            systemError(
              "DeployZipPackageError",
              "Deployment accepted but no Location header returned",
              { source }
            )
          );
        }
        return ok(location);
      }

      return err(
        userError("DeployZipPackageError", `Zip deploy returned status ${res.status}`, { source })
      );
    } catch (e) {
      lastError = e;
      if (axios.isAxiosError(e) && e.response && e.response.status >= 500 && attempt < maxRetries) {
        // Retry on 5xx
        continue;
      }
      break;
    }
  }

  return err(
    systemError(
      "DeployZipPackageError",
      `Failed to upload ZIP: ${lastError instanceof Error ? lastError.message : String(lastError)}`,
      { source }
    )
  );
}

/**
 * Poll the deployment status URL until complete.
 */
async function checkDeployStatus(
  locationUrl: string,
  token: string,
  source: string
): Promise<Result<void, AtkError>> {
  const headers = {
    Authorization: `Bearer ${token}`,
  };

  for (let i = 0; i < DEPLOY_CHECK_MAX_ATTEMPTS; i++) {
    await new Promise((resolve) => setTimeout(resolve, DEPLOY_CHECK_INTERVAL_MS));

    try {
      const res = await axios.get<{ status?: number; message?: string }>(locationUrl, {
        headers,
        timeout: 30_000,
      });

      if (res.status === 202) {
        continue; // Still deploying
      }

      if (res.status === 200 || res.status === 201) {
        if (res.data?.status === DeployStatus.Failed) {
          return err(
            userError(
              "DeploymentFailed",
              `Deployment failed: ${res.data.message ?? "unknown reason"}`,
              { source }
            )
          );
        }
        return ok(undefined);
      }

      return err(
        systemError(
          "CheckDeploymentStatusError",
          `Unexpected status ${res.status} checking deployment`,
          { source }
        )
      );
    } catch (e) {
      return err(
        systemError(
          "CheckDeploymentStatusError",
          `Error checking deployment status: ${e instanceof Error ? e.message : String(e)}`,
          { source }
        )
      );
    }
  }

  return err(
    systemError("CheckDeploymentStatusTimeout", "Deployment status check timed out", { source })
  );
}
