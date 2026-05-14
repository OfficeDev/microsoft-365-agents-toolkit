// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { Result, ok, err } from "neverthrow";
import type { AtkError } from "../core/error";
import { userError, systemError } from "../core/error";

/** Default environment folder relative to project root. */
const ENV_DIR = "env";
/** Prefix for environment files. */
const ENV_PREFIX = ".env.";

/**
 * List all environment names in a project.
 * Reads folder `env/` and parses `.env.{name}` file names.
 */
export async function listEnvironments(projectPath: string): Promise<Result<string[], AtkError>> {
  const envDir = path.join(projectPath, ENV_DIR);
  try {
    const files = await fs.promises.readdir(envDir);
    const envNames = files
      .filter((f) => f.startsWith(ENV_PREFIX))
      .map((f) => f.slice(ENV_PREFIX.length))
      .filter((name) => name.length > 0);
    return ok(envNames);
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") {
      return ok([]); // No env folder yet
    }
    return err(
      systemError("ListEnvironmentsFailed", `Failed to list environments: ${e}`, {
        source: "environment",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Read all variables from an environment file.
 */
export async function readEnvFile(
  projectPath: string,
  envName: string
): Promise<Result<Record<string, string>, AtkError>> {
  const envFilePath = path.join(projectPath, ENV_DIR, `${ENV_PREFIX}${envName}`);
  try {
    const content = await fs.promises.readFile(envFilePath, "utf-8");
    return ok(parseEnvContent(content));
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") {
      return err(
        userError("EnvironmentNotFound", `Environment "${envName}" not found.`, {
          source: "environment",
        })
      );
    }
    return err(
      systemError("ReadEnvFailed", `Failed to read env file: ${e}`, {
        source: "environment",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Write variables to an environment file, creating the env/ directory if needed.
 */
export async function writeEnvFile(
  projectPath: string,
  envName: string,
  variables: Record<string, string>
): Promise<Result<void, AtkError>> {
  const envDir = path.join(projectPath, ENV_DIR);
  const envFilePath = path.join(envDir, `${ENV_PREFIX}${envName}`);
  try {
    await fs.promises.mkdir(envDir, { recursive: true });
    const content = serializeEnvContent(variables);
    await fs.promises.writeFile(envFilePath, content, "utf-8");
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("WriteEnvFailed", `Failed to write env file: ${e}`, {
        source: "environment",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Add a new environment by creating its .env file.
 */
export async function addEnvironment(
  projectPath: string,
  envName: string,
  copyFrom?: string
): Promise<Result<void, AtkError>> {
  // Check it doesn't already exist
  const envs = await listEnvironments(projectPath);
  if (envs.isErr()) return err(envs.error);
  if (envs.value.includes(envName)) {
    return err(
      userError("EnvironmentExists", `Environment "${envName}" already exists.`, {
        source: "environment",
      })
    );
  }

  // Copy from another env or start empty
  let variables: Record<string, string> = {};
  if (copyFrom) {
    const source = await readEnvFile(projectPath, copyFrom);
    if (source.isErr()) return err(source.error);
    variables = source.value;
  }

  return writeEnvFile(projectPath, envName, variables);
}

/**
 * Reset (clear) an environment file.
 */
export async function resetEnvironment(
  projectPath: string,
  envName: string
): Promise<Result<void, AtkError>> {
  return writeEnvFile(projectPath, envName, {});
}

/**
 * Parse a `.env` file into key-value pairs.
 * Supports `KEY=VALUE`, `KEY="VALUE"`, comments (#), and empty lines.
 */
function parseEnvContent(content: string): Record<string, string> {
  const result: Record<string, string> = {};
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eqIndex = trimmed.indexOf("=");
    if (eqIndex === -1) continue;
    const key = trimmed.slice(0, eqIndex).trim();
    let value = trimmed.slice(eqIndex + 1).trim();
    // Strip surrounding quotes
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    result[key] = value;
  }
  return result;
}

/**
 * Serialize key-value pairs back into `.env` file format.
 */
function serializeEnvContent(variables: Record<string, string>): string {
  return (
    Object.entries(variables)
      .map(([key, value]) => `${key}=${value}`)
      .join("\n") + "\n"
  );
}
