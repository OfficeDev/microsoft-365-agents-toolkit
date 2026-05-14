// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { defineOperation } from "../core/operation";
import type { AtkContext } from "../core/context";
import { userError } from "../core/error";
import { readEnvFile, writeEnvFile } from "../environment/envManager";
import { parseProjectYaml } from "./parser";
import { executeLifecycle } from "./executor";
import { analyzeSteps } from "./analyze";
import { createProgressAdapter } from "./progress";
import {
  ensureM365Auth,
  ensureAzureAuth,
  ensureSubscription,
  ensureResourceGroup,
  confirmProvision,
  confirmDeploy,
  ensureResourceSuffix,
} from "./prerequisites";
import type {
  LifecycleProgress,
  LifecycleOperationResult,
  PostAction,
  LifecycleName,
  DriverStep,
} from "./types";

const SOURCE = "lifecycle/operations";
const YAML_FILE = "m365agents.yml";

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const provisionInputSchema = z.object({
  projectPath: z.string().min(1),
  envName: z.string().min(1),
  skipConsent: z.boolean().optional(),
});

const deployInputSchema = z.object({
  projectPath: z.string().min(1),
  envName: z.string().min(1),
  skipConsent: z.boolean().optional(),
});

const publishInputSchema = z.object({
  projectPath: z.string().min(1),
  envName: z.string().min(1),
});

type ProvisionInput = z.infer<typeof provisionInputSchema>;
type DeployInput = z.infer<typeof deployInputSchema>;
type PublishInput = z.infer<typeof publishInputSchema>;

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

async function loadYaml(projectPath: string) {
  const yamlPath = path.join(projectPath, YAML_FILE);
  try {
    const content = await fs.promises.readFile(yamlPath, "utf-8");
    return parseProjectYaml(content);
  } catch (_e) {
    return err(
      userError("YamlNotFound", `Could not read "${YAML_FILE}" in "${projectPath}".`, {
        source: SOURCE,
      })
    );
  }
}

async function loadEnvMap(projectPath: string, envName: string): Promise<Map<string, string>> {
  const result = await readEnvFile(projectPath, envName);
  if (result.isOk()) {
    return new Map(Object.entries(result.value));
  }
  // Env file may not exist yet — start with empty map
  return new Map();
}

async function persistEnvMap(
  projectPath: string,
  envName: string,
  envMap: Map<string, string>
): Promise<void> {
  const vars: Record<string, string> = {};
  for (const [k, v] of envMap) {
    vars[k] = v;
  }
  await writeEnvFile(projectPath, envName, vars);
}

function getProgress(ctx: AtkContext, progress?: LifecycleProgress): LifecycleProgress {
  return progress ?? createProgressAdapter(ctx.ui);
}

// ---------------------------------------------------------------------------
// provisionOp
// ---------------------------------------------------------------------------

/**
 * Provision cloud resources for the current project.
 *
 * Pipeline:
 *   1. Load environment file
 *   2. Parse m365agents.yml
 *   3. Analyze steps (M365? Azure?)
 *   4. Auth gates (M365 login, Azure login)
 *   5. Subscription & resource group selection
 *   6. Consent dialog
 *   7. Execute lifecycle
 *   8. Persist environment outputs
 */
export const provisionOp = defineOperation(
  "provision",
  provisionInputSchema,
  async (ctx: AtkContext, input: ProvisionInput) => {
    const { projectPath, envName, skipConsent } = input;
    const progress = getProgress(ctx);

    // 1. Load env
    const envMap = await loadEnvMap(projectPath, envName);

    // 2. Parse YAML
    const modelResult = await loadYaml(projectPath);
    if (modelResult.isErr()) return err(modelResult.error);
    const model = modelResult.value;

    const steps = getLifecycleSteps(model as unknown as Record<string, unknown>, "provision");
    if (!steps || steps.length === 0) {
      return err(
        userError("NoProvisionSteps", `No "provision" steps found in ${YAML_FILE}.`, {
          source: SOURCE,
        })
      );
    }

    // 3. Analyze
    const analysis = analyzeSteps(steps, envMap);

    // 4. Auth gates
    let m365Info;
    if (analysis.needsM365) {
      const m365Result = await ensureM365Auth(ctx);
      if (m365Result.isErr()) return err(m365Result.error);
      m365Info = m365Result.value;
    }

    let subInfo;
    if (analysis.needsAzure) {
      const azureResult = await ensureAzureAuth(ctx);
      if (azureResult.isErr()) return err(azureResult.error);

      // 5. Subscription & resource group
      ensureResourceSuffix(envMap);

      const subResult = await ensureSubscription(ctx, envMap);
      if (subResult.isErr()) return err(subResult.error);
      subInfo = subResult.value;

      if (
        analysis.unresolvedVars.includes("AZURE_RESOURCE_GROUP_NAME") ||
        !envMap.get("AZURE_RESOURCE_GROUP_NAME")
      ) {
        const projectName = path.basename(projectPath);
        const rgResult = await ensureResourceGroup(
          ctx,
          envMap,
          subInfo.subscriptionId,
          projectName,
          envName
        );
        if (rgResult.isErr()) return err(rgResult.error);
      }
    }

    // 6. Consent
    if (!skipConsent) {
      const consentResult = await confirmProvision(ctx, envName, m365Info, subInfo);
      if (consentResult.isErr()) return err(consentResult.error);
    }

    // 7. Execute
    const execResult = await executeLifecycle(ctx, "provision", steps, envMap, progress);
    if (execResult.isErr()) {
      // Persist partial outputs on failure
      await persistEnvMap(projectPath, envName, envMap);
      return err(execResult.error);
    }

    // 8. Persist
    await persistEnvMap(projectPath, envName, envMap);

    // Post-actions
    const postActions: PostAction[] = [
      {
        type: "showMessage",
        message: `Provision completed (${execResult.value.steps.length} steps).`,
      },
    ];
    if (subInfo) {
      const rgName = envMap.get("AZURE_RESOURCE_GROUP_NAME");
      if (rgName) {
        postActions.push({
          type: "openUrl",
          message: "View resources in Azure portal",
          url: `https://portal.azure.com/#@${subInfo.tenantId}/resource/subscriptions/${subInfo.subscriptionId}/resourceGroups/${rgName}`,
        });
      }
    }

    return ok({
      lifecycleResult: execResult.value,
      envMap,
      postActions,
    } satisfies LifecycleOperationResult);
  }
);

// ---------------------------------------------------------------------------
// deployOp
// ---------------------------------------------------------------------------

/**
 * Deploy artifacts for the current project.
 *
 * Pipeline:
 *   1. Load environment file
 *   2. Parse m365agents.yml
 *   3. Consent dialog (skipped for local envs)
 *   4. Execute lifecycle
 *   5. Persist environment outputs
 */
export const deployOp = defineOperation(
  "deploy",
  deployInputSchema,
  async (ctx: AtkContext, input: DeployInput) => {
    const { projectPath, envName, skipConsent } = input;
    const progress = getProgress(ctx);

    // 1. Load env
    const envMap = await loadEnvMap(projectPath, envName);

    // 2. Parse YAML
    const modelResult = await loadYaml(projectPath);
    if (modelResult.isErr()) return err(modelResult.error);
    const model = modelResult.value;

    const steps = getLifecycleSteps(model as unknown as Record<string, unknown>, "deploy");
    if (!steps || steps.length === 0) {
      return err(
        userError("NoDeploySteps", `No "deploy" steps found in ${YAML_FILE}.`, {
          source: SOURCE,
        })
      );
    }

    // 3. Consent
    if (!skipConsent) {
      const consentResult = await confirmDeploy(ctx, envName);
      if (consentResult.isErr()) return err(consentResult.error);
    }

    // 4. Execute
    const execResult = await executeLifecycle(ctx, "deploy", steps, envMap, progress);
    if (execResult.isErr()) {
      await persistEnvMap(projectPath, envName, envMap);
      return err(execResult.error);
    }

    // 5. Persist
    await persistEnvMap(projectPath, envName, envMap);

    return ok({
      lifecycleResult: execResult.value,
      envMap,
      postActions: [
        {
          type: "showMessage",
          message: `Deploy completed (${execResult.value.steps.length} steps).`,
        },
      ],
    } satisfies LifecycleOperationResult);
  }
);

// ---------------------------------------------------------------------------
// publishOp
// ---------------------------------------------------------------------------

/**
 * Publish the application to the org app catalog.
 *
 * Pipeline:
 *   1. Load environment file
 *   2. Parse m365agents.yml
 *   3. Execute lifecycle
 *   4. Persist environment outputs
 */
export const publishOp = defineOperation(
  "publish",
  publishInputSchema,
  async (ctx: AtkContext, input: PublishInput) => {
    const { projectPath, envName } = input;
    const progress = getProgress(ctx);

    // 1. Load env
    const envMap = await loadEnvMap(projectPath, envName);

    // 2. Parse YAML
    const modelResult = await loadYaml(projectPath);
    if (modelResult.isErr()) return err(modelResult.error);
    const model = modelResult.value;

    const steps = getLifecycleSteps(model as unknown as Record<string, unknown>, "publish");
    if (!steps || steps.length === 0) {
      return err(
        userError("NoPublishSteps", `No "publish" steps found in ${YAML_FILE}.`, {
          source: SOURCE,
        })
      );
    }

    // 3. Execute
    const execResult = await executeLifecycle(ctx, "publish", steps, envMap, progress);
    if (execResult.isErr()) {
      await persistEnvMap(projectPath, envName, envMap);
      return err(execResult.error);
    }

    // 4. Persist
    await persistEnvMap(projectPath, envName, envMap);

    return ok({
      lifecycleResult: execResult.value,
      envMap,
      postActions: [
        {
          type: "showMessage",
          message: `Publish completed (${execResult.value.steps.length} steps).`,
        },
        {
          type: "openUrl",
          message: "Manage app in Teams Admin Center",
          url: "https://admin.teams.microsoft.com/policies/manage-apps",
        },
      ],
    } satisfies LifecycleOperationResult);
  }
);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getLifecycleSteps(
  model: Record<string, unknown>,
  lifecycle: LifecycleName
): DriverStep[] | undefined {
  return model[lifecycle] as DriverStep[] | undefined;
}
