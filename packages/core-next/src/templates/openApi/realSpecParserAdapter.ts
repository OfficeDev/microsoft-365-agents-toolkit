// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Real SpecParserAdapter implementation backed by the inline spec parser.
 *
 * Replaces StubSpecParserAdapter with actual OpenAPI parsing, validation,
 * operation listing, and code generation.
 */

import type { OpenAPIV3 } from "openapi-types";
import type {
  SpecParserAdapter,
  SpecValidationResult as AdapterValidationResult,
  ApiOperationInfo,
  SpecGenerationResult,
} from "./specParserAdapter";
import { parseSpec } from "../../specParser/parser";
import { createValidator } from "../../specParser/validator";
import { filterSpec } from "../../specParser/filter";
import type { ParseOptions } from "../../specParser/types";
import { DEFAULT_PARSE_OPTIONS, ProjectType } from "../../specParser/types";
import { getAuthArray } from "../../specParser/utils";

/**
 * Real adapter that delegates to the inline spec parser modules.
 */
export class RealSpecParserAdapter implements SpecParserAdapter {
  private readonly options: ParseOptions;

  constructor(options?: Partial<ParseOptions>) {
    this.options = { ...DEFAULT_PARSE_OPTIONS, ...options };
  }

  async validate(specPath: string): Promise<AdapterValidationResult> {
    const parseResult = await parseSpec(specPath, this.options);
    if (parseResult.isErr()) {
      return {
        valid: false,
        errors: [parseResult.error.message],
        warnings: [],
      };
    }

    const { spec, warnings: parseWarnings } = parseResult.value;
    const validator = createValidator(spec.resolved, this.options);
    const specResult = validator.validateSpec();

    return {
      valid: specResult.errors.length === 0,
      errors: specResult.errors.map((e: { content: string }) => e.content),
      warnings: [
        ...parseWarnings.map((w: { content: string }) => w.content),
        ...specResult.warnings.map((w: { content: string }) => w.content),
      ],
    };
  }

  async listOperations(specPath: string): Promise<ApiOperationInfo[]> {
    const parseResult = await parseSpec(specPath, this.options);
    if (parseResult.isErr()) return [];

    const { spec } = parseResult.value;
    const validator = createValidator(spec.resolved, this.options);
    const apiMap = validator.listAPIs();

    const operations: ApiOperationInfo[] = [];
    for (const [key, mapValue] of Object.entries(apiMap)) {
      if (!mapValue.isValid) continue;

      const [method, ...pathParts] = key.split(" ");
      const path = pathParts.join(" ");
      const op = mapValue.operation;

      // Get auth info for this operation
      const authArray = getAuthArray(op.security, spec.resolved);
      let auth: ApiOperationInfo["auth"];
      if (authArray.length > 0 && authArray[0].length > 0) {
        const authScheme = authArray[0][0];
        auth = {
          authName: authScheme.name,
          authType: mapAuthType(authScheme.authScheme as OpenAPIV3.SecuritySchemeObject),
        };
      }

      operations.push({
        id: `${method} ${path}`,
        method: method.toUpperCase(),
        path,
        summary: op.summary,
        group: op.tags?.[0],
        auth,
      });
    }

    return operations;
  }

  async generate(
    specPath: string,
    operations: string[],
    _outputDir: string,
    projectType: "Copilot" | "TeamsAi" | "SME"
  ): Promise<SpecGenerationResult> {
    const opts: ParseOptions = {
      ...this.options,
      projectType: toProjectType(projectType),
    };

    const parseResult = await parseSpec(specPath, opts);
    if (parseResult.isErr()) {
      return {
        files: new Map(),
        warnings: [`Parse error: ${parseResult.error.message}`],
      };
    }

    const { spec, warnings: parseWarnings } = parseResult.value;

    // Filter spec to selected operations
    const filteredSpec = filterSpec(operations, spec.unresolved, spec.resolved, opts);

    // Return the filtered spec as the generated output
    // Full generator support (adaptive cards, plugin manifests) will be added in Phase 3
    const files = new Map<string, string>();
    files.set("openapi.json", JSON.stringify(filteredSpec, null, 2));

    return {
      files,
      warnings: parseWarnings.map((w: { content: string }) => w.content),
    };
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function toProjectType(type: "Copilot" | "TeamsAi" | "SME"): ProjectType {
  switch (type) {
    case "Copilot":
      return ProjectType.Copilot;
    case "TeamsAi":
      return ProjectType.TeamsAi;
    case "SME":
    default:
      return ProjectType.SME;
  }
}

function mapAuthType(
  scheme: OpenAPIV3.SecuritySchemeObject
): "apiKey" | "oauth2" | "bearer" | "microsoftEntra" | undefined {
  if (scheme.type === "apiKey") return "apiKey";
  if (scheme.type === "oauth2") return "oauth2";
  if (scheme.type === "http" && (scheme as OpenAPIV3.HttpSecurityScheme).scheme === "bearer") {
    return "bearer";
  }
  return undefined;
}
