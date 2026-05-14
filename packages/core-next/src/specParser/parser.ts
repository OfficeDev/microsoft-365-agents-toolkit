// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * OpenAPI spec parser — loads, validates, and dereferences OpenAPI documents.
 *
 * Wraps @apidevtools/swagger-parser for parsing/validation and
 * swagger2openapi for Swagger 2.0 → OpenAPI 3.0 conversion.
 *
 * All public functions return Result<T, AtkError> — never throw.
 */

import * as crypto from "crypto";
import SwaggerParser from "@apidevtools/swagger-parser";
import type { OpenAPIV3 } from "openapi-types";
import * as converter from "swagger2openapi";
import { ok, err, type Result } from "neverthrow";
import { userError, systemError, type AtkError } from "../core/error";
import type { ParsedSpec, ParseOptions, ValidationWarning } from "./types";
import { DEFAULT_PARSE_OPTIONS, WarningType } from "./types";

import { SpecParserMessages } from "./constants";

const SOURCE = "specParser/parser";

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse an OpenAPI specification file or document object.
 *
 * Pipeline:
 * 1. Parse raw file (YAML/JSON)
 * 2. Detect Swagger 2.0 → convert to OpenAPI 3.0 if allowed
 * 3. Clone and dereference (resolve all $ref pointers)
 * 4. Return both unresolved and resolved documents
 *
 * @param pathOrDoc - File path, URL, or pre-parsed OpenAPI document
 * @param options - Parse options (merged with defaults)
 */
export async function parseSpec(
  pathOrDoc: string | OpenAPIV3.Document,
  options?: ParseOptions
): Promise<Result<{ spec: ParsedSpec; warnings: ValidationWarning[] }, AtkError>> {
  const opts: Required<ParseOptions> = { ...DEFAULT_PARSE_OPTIONS, ...options };
  const warnings: ValidationWarning[] = [];

  try {
    const parser = new SwaggerParser();

    // Step 1: Parse (YAML/JSON → JS object), do NOT dereference yet
    let rawSpec = (await parser.parse(pathOrDoc as string)) as unknown as OpenAPIV3.Document;

    // Compute hash of original content
    const specHash = computeHash(rawSpec);

    // Step 2: Detect Swagger 2.0 and convert if allowed
    const isSwagger =
      !rawSpec.openapi && (rawSpec as unknown as Record<string, unknown>).swagger === "2.0";
    let isConverted = false;

    if (isSwagger) {
      if (!opts.allowSwagger) {
        return err(
          userError("SwaggerNotSupported", SpecParserMessages.SwaggerNotSupported, {
            source: SOURCE,
          })
        );
      }
      const converted = await convertSwagger2ToOpenApi3(rawSpec);
      if (converted.isErr()) return err(converted.error);
      rawSpec = converted.value;
      isConverted = true;
      warnings.push({
        type: WarningType.ConvertSwaggerToOpenAPI,
        content: SpecParserMessages.ConvertSwaggerToOpenAPI,
      });
    }

    // Step 3: Validate OpenAPI version
    if (rawSpec.openapi && !rawSpec.openapi.startsWith("3.0")) {
      return err(
        userError(
          "SpecVersionNotSupported",
          SpecParserMessages.SpecVersionNotSupported.replace("%s", rawSpec.openapi),
          { source: SOURCE }
        )
      );
    }

    // Step 4: Clone and dereference
    const cloned = JSON.parse(JSON.stringify(rawSpec)) as OpenAPIV3.Document;
    const resolved = (await new SwaggerParser().dereference(cloned)) as OpenAPIV3.Document;

    return ok({
      spec: {
        unresolved: rawSpec,
        resolved,
        isConverted,
        specHash,
      },
      warnings,
    });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return err(
      systemError("ParseSpecFailed", `Failed to parse OpenAPI spec: ${message}`, {
        source: SOURCE,
        inner: e instanceof Error ? e : undefined,
      })
    );
  }
}

// ---------------------------------------------------------------------------
// Swagger 2.0 → OpenAPI 3.0 conversion (isolated for replacement)
// ---------------------------------------------------------------------------

/**
 * Convert a Swagger 2.0 document to OpenAPI 3.0.
 * Isolated in a single function so that if swagger2openapi is ever
 * replaced, only this function needs to change.
 */
async function convertSwagger2ToOpenApi3(
  swaggerDoc: OpenAPIV3.Document
): Promise<Result<OpenAPIV3.Document, AtkError>> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await converter.convertObj(swaggerDoc as any, {} as any);
    return ok(
      (result as unknown as Record<string, unknown>).openapi as unknown as OpenAPIV3.Document
    );
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return err(
      userError(
        "SwaggerConversionFailed",
        `Failed to convert Swagger 2.0 to OpenAPI 3.0: ${message}`,
        {
          source: SOURCE,
          inner: e instanceof Error ? e : undefined,
        }
      )
    );
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function computeHash(doc: OpenAPIV3.Document): string {
  const content = JSON.stringify(doc);
  return crypto.createHash("sha256").update(content).digest("hex");
}

/**
 * Resolve environment variable placeholders in a spec.
 * Replaces `${{ENV_VAR}}` patterns with process.env values.
 */
export function resolveEnvVars(spec: OpenAPIV3.Document): OpenAPIV3.Document {
  let text = JSON.stringify(spec);
  const envPattern = /\$\{\{([^}]+)\}\}/g;
  text = text.replace(envPattern, (_match, envName: string) => {
    return process.env[envName.trim()] ?? `\${{${envName}}}`;
  });
  return JSON.parse(text) as OpenAPIV3.Document;
}

/**
 * Check if a spec has circular references.
 */
export async function hasCircularRefs(pathOrDoc: string | OpenAPIV3.Document): Promise<boolean> {
  try {
    const parser = new SwaggerParser();
    await parser.dereference(pathOrDoc as string);
    return parser.$refs.circular;
  } catch {
    return false;
  }
}
