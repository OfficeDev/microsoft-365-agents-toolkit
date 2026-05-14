// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { userError, systemError } from "../../core/error";
import type { AddExistingPluginInput, AddExistingPluginResult } from "../types";
import { findPlaceholders } from "../manifest/resolver";

/**
 * Add an existing plugin to a declarative agent manifest.
 *
 * 1. Reads & validates the source plugin manifest.
 * 2. Copies the OpenAPI spec into the project (renames on conflict).
 * 3. Saves the plugin manifest next to the agent manifest.
 * 4. Registers the action in the agent manifest.
 * 5. Collects warnings for unresolved ${{VAR}} placeholders.
 */
export async function addExistingPlugin(
  input: AddExistingPluginInput
): Promise<Result<AddExistingPluginResult, AtkError>> {
  const { agentManifestPath, pluginManifestPath, apiSpecPath, actionId } = input;

  try {
    // 1. Read and validate source plugin manifest
    const rawContent = await fs.promises.readFile(pluginManifestPath, "utf-8");
    const pluginManifest = JSON.parse(rawContent);

    const validationError = validatePluginManifest(pluginManifest);
    if (validationError) {
      return err(validationError);
    }

    const outputDir = path.dirname(agentManifestPath);
    const runtimes = pluginManifest.runtimes as Array<{
      type: string;
      spec?: { url: string };
    }>;
    const openApiRuntime = runtimes.find((r) => r.type === "OpenApi");
    const destRelativePath = openApiRuntime!.spec!.url;

    // 2. Copy OpenAPI spec (avoid overwriting existing files)
    const specDestFullPath = await copySpecFile(apiSpecPath, outputDir, destRelativePath);
    const specRelative = path.relative(outputDir, specDestFullPath).replace(/\\/g, "/");

    // Update runtime spec paths if they changed
    for (const runtime of runtimes) {
      if (runtime.type === "OpenApi" && runtime.spec?.url) {
        runtime.spec.url = specRelative;
      }
    }

    // 3. Write plugin manifest to next available filename
    const destPluginPath = nextAvailablePath(path.join(outputDir, "ai-plugin.json"), outputDir);
    await fs.promises.writeFile(destPluginPath, JSON.stringify(pluginManifest, null, 4), "utf-8");

    // 4. Register in the DA manifest via the wrapper
    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);
    wrapper.addAction(actionId, path.relative(outputDir, destPluginPath).replace(/\\/g, "/"));
    await wrapper.save(agentManifestPath);

    // 5. Collect warnings for unresolved placeholders
    const warnings: string[] = [];

    const pluginPlaceholders = findPlaceholders(JSON.stringify(pluginManifest));
    if (pluginPlaceholders.length > 0) {
      warnings.push(
        `Plugin manifest references environment variables: ${pluginPlaceholders.map((p) => p.name).join(", ")}`
      );
    }

    try {
      const specContent = await fs.promises.readFile(specDestFullPath, "utf-8");
      const specPlaceholders = findPlaceholders(specContent);
      if (specPlaceholders.length > 0) {
        warnings.push(
          `API spec references environment variables: ${specPlaceholders.map((p) => p.name).join(", ")}`
        );
      }
    } catch {
      // Best-effort — don't fail the operation for spec-read issues
    }

    return ok({ warnings, destinationPluginManifestPath: destPluginPath });
  } catch (e) {
    return err(
      systemError("AddExistingPluginFailed", `Failed to add existing plugin: ${e}`, {
        source: "declarativeAgent/actions",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Validate that a plugin manifest has the required structure.
 */
function validatePluginManifest(manifest: Record<string, unknown>): AtkError | undefined {
  if (!manifest.schema_version) {
    return userError("InvalidPluginManifest", "Plugin manifest is missing schema_version.", {
      source: "declarativeAgent/actions",
    });
  }
  const runtimes = manifest.runtimes as Array<{ type: string; spec?: { url: string } }>;
  if (!Array.isArray(runtimes) || runtimes.length === 0) {
    return userError("InvalidPluginManifest", "Plugin manifest must have at least one runtime.", {
      source: "declarativeAgent/actions",
    });
  }
  const openApiRuntime = runtimes.find((r) => r.type === "OpenApi");
  if (!openApiRuntime?.spec?.url) {
    return userError(
      "InvalidPluginManifest",
      "Plugin manifest must have an OpenApi runtime with a spec URL.",
      { source: "declarativeAgent/actions" }
    );
  }
  return undefined;
}

/**
 * Copy a spec file to the output dir, returning the destination absolute path.
 * If the target already exists, append a numeric suffix.
 */
async function copySpecFile(
  sourcePath: string,
  outputDir: string,
  relativeTarget: string
): Promise<string> {
  let destPath = path.resolve(outputDir, relativeTarget);

  // If file exists or path escapes the output dir, use apiSpecFolder with unique name
  const cross = path.relative(outputDir, destPath).startsWith("..");
  if (cross || existsSync(destPath)) {
    destPath = nextAvailablePath(
      path.join(outputDir, "apiSpecificationFile", path.basename(sourcePath)),
      outputDir
    );
  }

  await fs.promises.mkdir(path.dirname(destPath), { recursive: true });
  await fs.promises.copyFile(sourcePath, destPath);
  return destPath;
}

/**
 * Find the next available file path by appending _N suffix before the extension.
 */
function nextAvailablePath(basePath: string, _outputDir: string): string {
  if (!existsSync(basePath)) {
    return basePath;
  }
  const ext = path.extname(basePath);
  const stem = basePath.slice(0, -ext.length);
  let i = 1;
  let candidate: string;
  do {
    candidate = `${stem}_${i}${ext}`;
    i++;
  } while (existsSync(candidate));
  return candidate;
}

function existsSync(p: string): boolean {
  try {
    fs.accessSync(p);
    return true;
  } catch {
    return false;
  }
}
