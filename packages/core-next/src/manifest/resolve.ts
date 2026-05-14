// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { ok, err, Result } from "neverthrow";
import type { AtkContext } from "../core/context";
import { userError, systemError, AtkError } from "../core/error";
import type { EnvResolutionResult, EnvPlaceholder, ManifestType } from "./types";

// ─── Regex patterns ──────────────────────────────────────────

/** Matches ${{VAR_NAME}} env-var placeholders. */
const PLACEHOLDER_RE = /\$\{\{([A-Za-z_][A-Za-z0-9_]*)\}\}/g;

/** Matches $[ funcName(...) ] function expressions (outermost brackets). */
const FUNCTION_EXPR_RE = /\$\[ *[a-zA-Z][a-zA-Z]*\([^\]]*\) *\]/g;

/** Allowed file extensions for $[file()] content inclusion. */
const ALLOWED_FILE_EXTENSIONS = new Set([".txt", ".md"]);

const SOURCE = "ManifestResolve";
const HELP_LINK = "https://aka.ms/teamsfx-customize-manifest";

// ─── Env-var placeholder resolution ──────────────────────────

/**
 * Resolve all `${{VAR}}` placeholders in `content` using the provided env map.
 *
 * Unlike the `declarativeAgent/manifest/resolver.ts` variant (which silently
 * leaves unresolved tokens), this function tracks every unresolved placeholder
 * so callers can enforce strict resolution when needed (e.g. packaging).
 */
export function resolveEnvPlaceholders(
  content: string,
  envs?: Readonly<Record<string, string>>
): EnvResolutionResult {
  const unresolved: EnvPlaceholder[] = [];

  const resolved = content.replace(PLACEHOLDER_RE, (full, varName: string) => {
    const value = envs?.[varName] ?? process.env[varName];
    if (value !== undefined && value !== "") {
      return value;
    }
    unresolved.push({ token: full, name: varName });
    return full;
  });

  return { content: resolved, unresolved };
}

/**
 * Extract the list of env-var names referenced in `content`.
 */
export function getEnvVariables(content: string): string[] {
  const matches = content.match(PLACEHOLDER_RE);
  if (!matches) return [];
  return [...new Set(matches.map((m) => m.slice(3, -2).trim()))];
}

// ─── $[file()] function expansion ────────────────────────────

/**
 * Expand `$[file('path')]` function expressions in manifest content.
 *
 * Supports:
 * - Static path:  `$[file('relative/path.txt')]`
 * - Env-var path: `$[file(${{VAR_NAME}})]`
 * - Nested:       `$[file(file(...))]`
 *
 * When `isJson` is true, the file content is JSON-escaped so it can
 * safely embed inside a JSON string value.
 *
 * @param content      The manifest string that may contain function expressions
 * @param ctx          Context for logging and telemetry
 * @param envs         Env-var map for resolving `${{VAR}}` inside function args
 * @param isJson       Whether to JSON-escape the file contents
 * @param manifestType The type of manifest being processed (for telemetry)
 * @param fromPath     The absolute path of the file containing the expression
 *                     (used to resolve relative paths)
 */
export async function expandFunctionExpressions(
  content: string,
  ctx: AtkContext,
  envs: Readonly<Record<string, string>> | undefined,
  isJson: boolean,
  manifestType: ManifestType,
  fromPath: string
): Promise<Result<string, AtkError>> {
  const matches = content.match(FUNCTION_EXPR_RE);
  if (!matches) return ok(content);

  let expandedCount = 0;

  for (const placeholder of matches) {
    // Strip outer $[ and ] then trim
    const inner = placeholder.slice(2, -1).trim();
    const result = await processFunction(inner, ctx, envs, fromPath);
    if (result.isErr()) return err(result.error);

    let value = result.value;
    if (isJson && value) {
      // JSON-encode the value and strip the wrapping quotes
      value = JSON.stringify(value).slice(1, -1);
    }
    if (value) {
      expandedCount += 1;
      content = content.replace(placeholder, value);
    }
  }

  if (expandedCount > 0) {
    ctx.telemetry.sendTelemetryEvent("manifest-with-function", {
      "manifest-type": manifestType,
      "function-count": String(expandedCount),
    });
  }

  return ok(content);
}

// ─── Internal: recursive function processor ──────────────────

async function processFunction(
  expr: string,
  ctx: AtkContext,
  envs: Readonly<Record<string, string>> | undefined,
  fromPath: string
): Promise<Result<string, AtkError>> {
  const trimmed = expr.trim();

  if (!trimmed.startsWith("file(") || !trimmed.endsWith(")")) {
    ctx.logger.error(`Unsupported manifest function: "${trimmed}". Only file() is supported.`);
    return err(
      userError(
        "UnsupportedManifestFunction",
        `Unsupported manifest function: "${trimmed}". Only file() is supported.`,
        {
          source: SOURCE,
          help: HELP_LINK,
        }
      )
    );
  }

  // Extract the parameter inside file(...)
  const param = trimmed.slice(5, -1).trim();

  if (param.startsWith("'") && param.endsWith("'")) {
    // Static string: file('relative/path.txt')
    const filePath = param.slice(1, -1);
    return readFunctionFile(filePath, ctx, envs, fromPath);
  }

  if (param.startsWith("${{") && param.endsWith("}}")) {
    // Env-var reference: file(${{VAR_NAME}})
    const { content: resolvedPath } = resolveEnvPlaceholders(param, envs);
    return readFunctionFile(resolvedPath, ctx, envs, fromPath);
  }

  if (param.startsWith("file(") && param.endsWith(")")) {
    // Nested function: file(file(...))
    const innerResult = await processFunction(param, ctx, envs, fromPath);
    if (innerResult.isErr()) return err(innerResult.error);
    return readFunctionFile(innerResult.value, ctx, envs, fromPath);
  }

  ctx.logger.error(`Invalid parameter for file() function: "${param}"`);
  return err(
    userError("InvalidFunctionParameter", `Invalid parameter for file() function: "${param}"`, {
      source: SOURCE,
      help: HELP_LINK,
    })
  );
}

async function readFunctionFile(
  filePath: string,
  ctx: AtkContext,
  envs: Readonly<Record<string, string>> | undefined,
  fromPath: string
): Promise<Result<string, AtkError>> {
  const ext = path.extname(filePath).toLowerCase();
  if (!ALLOWED_FILE_EXTENSIONS.has(ext)) {
    ctx.logger.error(
      `Unsupported file type "${ext}" in file() function. Only .txt and .md are allowed.`
    );
    return err(
      userError(
        "UnsupportedFileFormat",
        `Unsupported file type "${ext}" in file() function. Only .txt and .md are allowed.`,
        {
          source: SOURCE,
          help: HELP_LINK,
        }
      )
    );
  }

  const absolutePath = path.isAbsolute(filePath)
    ? filePath
    : path.join(path.dirname(fromPath), filePath);

  try {
    let fileContent = await fs.promises.readFile(absolutePath, "utf8");
    // Strip BOM if present
    if (fileContent.charCodeAt(0) === 0xfeff) {
      fileContent = fileContent.slice(1);
    }
    // Resolve any env-var placeholders inside the included file
    const { content: resolved } = resolveEnvPlaceholders(fileContent, envs);
    // Normalize line endings
    return ok(resolved.replace(/\r\n/g, "\n"));
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") {
      return err(
        userError("FileNotFound", `File referenced in manifest function not found: ${filePath}`, {
          source: SOURCE,
          help: HELP_LINK,
        })
      );
    }
    ctx.logger.error(`Failed to read file "${absolutePath}": ${e}`);
    return err(
      systemError(
        "ReadFileError",
        `Failed to read file referenced in manifest function: ${absolutePath}`,
        {
          source: SOURCE,
          inner: e instanceof Error ? e : new Error(String(e)),
        }
      )
    );
  }
}

// ─── Orchestrator ────────────────────────────────────────────

/**
 * Fully resolve a manifest string:
 * 1. Expand `$[file()]` function expressions
 * 2. Resolve `${{VAR}}` env-var placeholders
 * 3. Optionally fail if any placeholders remain unresolved
 *
 * @param content       The raw manifest template string
 * @param ctx           Context for logging and telemetry
 * @param options       Resolution options
 * @returns The fully resolved content string, or an AtkError
 */
export async function resolveManifest(
  content: string,
  ctx: AtkContext,
  options?: {
    envs?: Readonly<Record<string, string>>;
    manifestType?: ManifestType;
    fromPath?: string;
    /** When true (default), returns an error if any ${{VAR}} remains unresolved. */
    strict?: boolean;
  }
): Promise<Result<string, AtkError>> {
  const { envs, manifestType, fromPath, strict = true } = options ?? {};

  // Step 1: expand $[file()] expressions (they may contain ${{VAR}} references internally)
  if (manifestType && fromPath) {
    const expandResult = await expandFunctionExpressions(
      content,
      ctx,
      envs,
      true, // manifest content is always JSON
      manifestType,
      fromPath
    );
    if (expandResult.isErr()) return expandResult;
    content = expandResult.value;
  }

  // Step 2: resolve ${{VAR}} placeholders
  const { content: resolved, unresolved } = resolveEnvPlaceholders(content, envs);

  // Step 3: strict mode — error on unresolved placeholders
  if (strict && unresolved.length > 0) {
    const missing = unresolved.map((u) => u.name).join(", ");
    return err(
      userError(
        "MissingEnvironmentVariables",
        `The following environment variables are not set: ${missing}. Set them in your .env file or environment.`,
        { source: SOURCE, help: HELP_LINK }
      )
    );
  }

  return ok(resolved);
}
