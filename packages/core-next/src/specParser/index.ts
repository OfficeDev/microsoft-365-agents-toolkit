// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * specParser barrel — public API surface for the spec parser module.
 */

// Types & enums
export {
  ParsedSpec,
  ValidationResult,
  ValidationStatus,
  ValidationError,
  ValidationWarning,
  ErrorType,
  WarningType,
  ProjectType,
  ParseOptions,
  DEFAULT_PARSE_OPTIONS,
  ListAPIInfo,
  ListAPIResult,
  APIValidationResult,
  APIMap,
  InvalidAPIInfo,
  AuthInfo,
  AuthType,
  OperationAuthInfoMap,
  AdaptiveCardBody,
  ArrayElement,
  ImageElement,
  TextBlockElement,
  TextElement,
  GenerateResult,
  CheckParamResult,
  ParamInfo,
} from "./types";

// Parser
export { parseSpec, resolveEnvVars, hasCircularRefs } from "./parser";

// Validator
export {
  Validator,
  CopilotValidator,
  SMEValidator,
  TeamsAIValidator,
  createValidator,
  type SpecValidationResult,
} from "./validator";

// Filter & optimizer
export { filterSpec } from "./filter";
export { optimizeSpec, type OptimizerOptions } from "./optimizer";

// Utilities
export {
  isBearerTokenAuth,
  isAPIKeyAuth,
  isAPIKeyAuthButNotInCookie,
  isOAuthWithAuthCodeFlow,
  isNotSupportedAuth,
  getAuthArray,
  getAuthMap,
  getAuthInfo,
  getResponseJson,
  getJsonContentType,
  convertPathToCamelCase,
  getSafeRegistrationIdEnvName,
  checkServerUrl,
  validateServer,
  getServerObject,
  isObjectSchema,
  isWellKnownName,
  formatStr,
  generateParametersFromSchema,
  getAuthSchemaObject,
} from "./utils";

// Constants
export {
  SpecParserMessages,
  HTTPMethods,
  ResponseCodeFor20X,
  WellKnownNames,
  Limits,
} from "./constants";
