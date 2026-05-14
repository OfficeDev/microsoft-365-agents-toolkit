// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Validation result from parsing an OpenAPI specification.
 */
export interface SpecValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * A single API operation parsed from an OpenAPI spec.
 */
export interface ApiOperationInfo {
  /** Unique operation identifier (e.g. "GET /pets") */
  id: string;
  /** HTTP method */
  method: string;
  /** API path */
  path: string;
  /** Human-readable summary */
  summary?: string;
  /** Group/tag name for UI grouping */
  group?: string;
  /** Authentication info for this operation */
  auth?: {
    authName?: string;
    authType?: "apiKey" | "oauth2" | "bearer" | "microsoftEntra";
  };
}

/**
 * Result of code generation from OpenAPI spec.
 */
export interface SpecGenerationResult {
  /** Generated source files: relative path → content */
  files: Map<string, string>;
  /** Warnings encountered during generation */
  warnings: string[];
}

/**
 * Abstraction over the spec parser, allowing a stub for testing
 * and a real implementation later (backed by @microsoft/m365-spec-parser).
 */
export interface SpecParserAdapter {
  /**
   * Validate an OpenAPI specification.
   * @param specPath - Path or URL to the OpenAPI spec file.
   */
  validate(specPath: string): Promise<SpecValidationResult>;

  /**
   * List available operations from an OpenAPI specification.
   * @param specPath - Path or URL to the OpenAPI spec file.
   */
  listOperations(specPath: string): Promise<ApiOperationInfo[]>;

  /**
   * Generate code artifacts from selected API operations.
   * @param specPath - Path or URL to the OpenAPI spec file.
   * @param operations - Selected operation IDs.
   * @param outputDir - Directory to write generated files into.
   * @param projectType - The project type context.
   */
  generate(
    specPath: string,
    operations: string[],
    outputDir: string,
    projectType: "Copilot" | "TeamsAi" | "SME"
  ): Promise<SpecGenerationResult>;
}

/**
 * Stub implementation of SpecParserAdapter.
 * Returns placeholder results for testing and development.
 * Replace with the real @microsoft/m365-spec-parser integration later.
 */
export class StubSpecParserAdapter implements SpecParserAdapter {
  validate(_specPath: string): Promise<SpecValidationResult> {
    return Promise.resolve({ valid: true, errors: [], warnings: [] });
  }

  listOperations(_specPath: string): Promise<ApiOperationInfo[]> {
    return Promise.resolve([
      {
        id: "GET /api/items",
        method: "GET",
        path: "/api/items",
        summary: "List items",
        group: "Items",
      },
      {
        id: "POST /api/items",
        method: "POST",
        path: "/api/items",
        summary: "Create item",
        group: "Items",
      },
    ]);
  }

  generate(
    _specPath: string,
    _operations: string[],
    _outputDir: string,
    _projectType: "Copilot" | "TeamsAi" | "SME"
  ): Promise<SpecGenerationResult> {
    return Promise.resolve({
      files: new Map(),
      warnings: ["Stub adapter: no real files generated"],
    });
  }
}

/**
 * Factory function to create a SpecParserAdapter.
 * Returns the real implementation backed by the inline spec parser.
 * Use StubSpecParserAdapter in tests for isolation.
 */
export function createSpecParserAdapter(): SpecParserAdapter {
  // Lazy import to avoid circular dependency at module scope
  const { RealSpecParserAdapter } = require("./realSpecParserAdapter");
  return new RealSpecParserAdapter();
}
