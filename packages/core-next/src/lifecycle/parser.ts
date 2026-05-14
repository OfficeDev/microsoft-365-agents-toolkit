// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import type { AtkError } from "../core/error";
import { userError, systemError } from "../core/error";
import type { RawProjectModel, DriverStep } from "./types";
import { LIFECYCLE_NAMES } from "./types";

/**
 * Parse a YAML string (m365agents.yml content) into a RawProjectModel.
 *
 * Uses the `js-yaml` or `yaml` library dynamically so consumers that
 * don't need lifecycle parsing don't pay the import cost.
 */
export async function parseProjectYaml(
  yamlContent: string
): Promise<Result<RawProjectModel, AtkError>> {
  try {
    // Dynamic import to keep the yaml dep optional at the module boundary
    const { load } = await import("js-yaml");
    const raw = load(yamlContent) as Record<string, unknown> | undefined;

    if (!raw || typeof raw !== "object") {
      return err(
        userError("InvalidYamlSchema", "YAML content is not a valid object.", {
          source: "lifecycle/parser",
        })
      );
    }

    return parseRawModel(raw);
  } catch (e) {
    return err(
      systemError("YamlParseError", `Failed to parse YAML: ${e}`, {
        source: "lifecycle/parser",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Parse a raw JS object (from YAML) into a validated RawProjectModel.
 */
function parseRawModel(obj: Record<string, unknown>): Result<RawProjectModel, AtkError> {
  // version — required
  if (!("version" in obj) || typeof obj.version !== "string") {
    return err(
      userError("MissingYamlField", 'YAML file must contain a "version" field of type string.', {
        source: "lifecycle/parser",
      })
    );
  }

  const model: RawProjectModel = { version: obj.version };

  // environmentFolderPath — optional string
  if ("environmentFolderPath" in obj) {
    if (typeof obj.environmentFolderPath !== "string") {
      return err(
        userError("InvalidYamlFieldType", '"environmentFolderPath" must be a string.', {
          source: "lifecycle/parser",
        })
      );
    }
    model.environmentFolderPath = obj.environmentFolderPath;
  }

  // additionalMetadata — optional object (no strict validation)
  if ("additionalMetadata" in obj && typeof obj.additionalMetadata === "object") {
    model.additionalMetadata = obj.additionalMetadata as Record<string, unknown>;
  }

  // Lifecycle sections
  for (const name of LIFECYCLE_NAMES) {
    if (name in obj) {
      const value = obj[name];
      if (!Array.isArray(value)) {
        return err(
          userError("InvalidYamlFieldType", `"${name}" must be an array.`, {
            source: "lifecycle/parser",
          })
        );
      }

      const stepsResult = parseSteps(name, value);
      if (stepsResult.isErr()) return err(stepsResult.error);
      model[name] = stepsResult.value;
    }
  }

  return ok(model);
}

/**
 * Parse and validate an array of driver step definitions.
 */
function parseSteps(lifecycleName: string, items: unknown[]): Result<DriverStep[], AtkError> {
  const steps: DriverStep[] = [];

  for (let i = 0; i < items.length; i++) {
    const item = items[i] as Record<string, unknown>;

    if (typeof item.uses !== "string") {
      return err(
        userError("MissingYamlField", `"${lifecycleName}[${i}].uses" must be a string.`, {
          source: "lifecycle/parser",
        })
      );
    }

    if (typeof item.with !== "object" || item.with === null) {
      return err(
        userError("MissingYamlField", `"${lifecycleName}[${i}].with" must be an object.`, {
          source: "lifecycle/parser",
        })
      );
    }

    const step: DriverStep = {
      uses: item.uses,
      with: item.with as Record<string, unknown>,
    };

    if (typeof item.name === "string") {
      step.name = item.name;
    }

    if (item.env && typeof item.env === "object" && !Array.isArray(item.env)) {
      step.env = item.env as Record<string, string>;
    }

    if (
      item.writeToEnvironmentFile &&
      typeof item.writeToEnvironmentFile === "object" &&
      !Array.isArray(item.writeToEnvironmentFile)
    ) {
      step.writeToEnvironmentFile = item.writeToEnvironmentFile as Record<string, string>;
    }

    steps.push(step);
  }

  return ok(steps);
}
