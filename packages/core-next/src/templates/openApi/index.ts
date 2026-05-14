// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type {
  SpecParserAdapter,
  SpecValidationResult,
  ApiOperationInfo,
  SpecGenerationResult,
} from "./specParserAdapter";
export { StubSpecParserAdapter, createSpecParserAdapter } from "./specParserAdapter";
export { RealSpecParserAdapter } from "./realSpecParserAdapter";
export { makeOpenApiScaffoldFn } from "./scaffoldFn";
export type { OpenApiProjectType } from "./scaffoldFn";
