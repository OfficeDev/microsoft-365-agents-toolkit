// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import type { AtkContext } from "../core/context";
import type { AtkError } from "../core/error";
import { userError } from "../core/error";
import { driverRegistry } from "../drivers/registry";
import type { DriverOutput } from "../drivers/types";
import type {
  DriverStep,
  LifecycleName,
  LifecycleResult,
  LifecycleProgress,
  StepResult,
} from "./types";
import { resolveConfig } from "./resolver";

/**
 * Execute a lifecycle (ordered list of driver steps).
 *
 * For each step:
 *   1. Look up the driver in the DriverRegistry.
 *   2. Resolve ${{VAR}} placeholders in the config.
 *   3. Execute the driver.
 *   4. Write outputs to the environment map.
 *   5. Report telemetry.
 *
 * @param ctx           The AtkContext for DI.
 * @param lifecycle     The lifecycle name being executed (for telemetry).
 * @param steps         The ordered driver steps to execute.
 * @param envMap        Mutable environment variable map; populated during execution.
 * @param progress      Optional progress callbacks for UI reporting.
 */
export async function executeLifecycle(
  ctx: AtkContext,
  lifecycle: LifecycleName,
  steps: DriverStep[],
  envMap: Map<string, string>,
  progress?: LifecycleProgress
): Promise<Result<LifecycleResult, AtkError>> {
  const stepResults: StepResult[] = [];
  const totalStart = Date.now();

  ctx.logger.info(`[lifecycle] Starting "${lifecycle}" with ${steps.length} step(s)`);
  await progress?.onStart(lifecycle, steps.length);

  for (let i = 0; i < steps.length; i++) {
    const step = steps[i];
    const driver = driverRegistry.get(step.uses);

    if (!driver) {
      return err(
        userError(
          "DriverNotFound",
          `Driver "${step.uses}" is not registered. Check m365agents.yml step ${i + 1} in "${lifecycle}".`,
          { source: "lifecycle/executor" }
        )
      );
    }

    // Inject step-level env vars
    if (step.env) {
      for (const [k, v] of Object.entries(step.env)) {
        // Resolve placeholders in env values too
        const { resolved } = resolveConfig({ val: v }, envMap);
        envMap.set(k, resolved.val as string);
      }
    }

    // Resolve placeholders in the driver config
    const { resolved: resolvedConfig, unresolved } = resolveConfig(step.with, envMap);

    // Auto-inject projectPath from context when the YAML step doesn't provide
    // it explicitly.  Many drivers (e.g. teamsApp/zipAppPackage) need a root
    // project path to resolve relative paths in their config.
    if (!resolvedConfig.projectPath && ctx.projectPath) {
      resolvedConfig.projectPath = ctx.projectPath;
    }

    if (unresolved.length > 0) {
      ctx.logger.warning(
        `[lifecycle] Step ${i + 1} ("${step.uses}") has unresolved variables: ${unresolved.map((u) => u.name).join(", ")}`
      );
    }

    const stepStart = Date.now();
    const stepLabel = step.name ? `${step.name} (${step.uses})` : step.uses;
    ctx.logger.info(`[lifecycle] Step ${i + 1}/${steps.length}: ${stepLabel}`);
    await progress?.onStepStart(i, stepLabel);

    // Temporarily sync envMap into process.env so that drivers which load
    // external files (e.g. ARM parameter JSON) can resolve ${{VAR}} placeholders
    // produced by earlier lifecycle steps.  Only set vars that are not already
    // present in process.env to avoid overwriting real environment values.
    const injectedKeys: string[] = [];
    for (const [k, v] of envMap) {
      if (process.env[k] === undefined) {
        process.env[k] = v;
        injectedKeys.push(k);
      }
    }

    let result: Result<DriverOutput, AtkError>;
    try {
      result = await driver.executeFn(ctx, resolvedConfig);
    } finally {
      // Clean up injected vars to avoid leaking state between steps
      for (const k of injectedKeys) {
        delete process.env[k];
      }
    }

    const durationMs = Date.now() - stepStart;

    if (result.isErr()) {
      ctx.logger.error(`[lifecycle] Step ${i + 1} failed: ${result.error.message}`);
      await progress?.onEnd(false);
      return err(result.error);
    }

    const outputs = result.value.outputs;

    // Write driver outputs to environment map.
    if (step.writeToEnvironmentFile && outputs) {
      // When writeToEnvironmentFile is specified, map camelCase YAML keys to
      // UPPER_CASE env-var names. Drivers may key their outputs by either form.
      for (const [yamlKey, envVarName] of Object.entries(step.writeToEnvironmentFile)) {
        const outputValue = outputs[yamlKey] ?? outputs[envVarName];
        if (outputValue !== undefined) {
          envMap.set(envVarName, outputValue);
        }
      }
    } else if (outputs) {
      // No writeToEnvironmentFile — write all outputs directly to envMap.
      // This is the common case for drivers like arm/deploy that produce
      // uppercase environment variable names as output keys.
      for (const [key, value] of Object.entries(outputs)) {
        if (value !== undefined) {
          envMap.set(key, value);
        }
      }
    }

    stepResults.push({ driver: step.uses, outputs, durationMs });
    ctx.logger.info(`[lifecycle] Step ${i + 1} completed (${durationMs}ms)`);
    await progress?.onStepComplete(i, stepLabel, durationMs);
  }

  const totalDurationMs = Date.now() - totalStart;
  ctx.logger.info(`[lifecycle] "${lifecycle}" completed (${totalDurationMs}ms)`);
  await progress?.onEnd(true);

  return ok({ lifecycle, steps: stepResults, totalDurationMs });
}
