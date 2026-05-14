// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Types for the spec-parser module.
 *
 * Redesigned from @microsoft/m365-spec-parser interfaces.ts:
 * - Enums use string values for JSON-safe serialization
 * - No class hierarchy — plain objects with discriminated fields
 * - Follows core-next error conventions (AtkError, Result pattern)
 */

import type { OpenAPIV3 } from "openapi-types";

// ---------------------------------------------------------------------------
// Parsed spec
// ---------------------------------------------------------------------------

/**
 * Result of parsing an OpenAPI specification file.
 * Contains both the original (with $refs) and the resolved (dereferenced) document.
 */
export interface ParsedSpec {
  /** The unresolved spec (preserves $ref pointers) */
  unresolved: OpenAPIV3.Document;
  /** The fully dereferenced spec */
  resolved: OpenAPIV3.Document;
  /** Whether the original was Swagger 2.0 (converted to 3.0) */
  isConverted: boolean;
  /** SHA-256 hash of the original file content */
  specHash: string;
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

export enum ValidationStatus {
  Valid = "valid",
  Warning = "warning",
  Error = "error",
}

export interface ValidationResult {
  status: ValidationStatus;
  warnings: ValidationWarning[];
  errors: ValidationError[];
  specHash?: string;
}

export interface ValidationWarning {
  type: WarningType;
  content: string;
  data?: unknown;
}

export interface ValidationError {
  type: ErrorType;
  content: string;
  data?: unknown;
}

/**
 * Error types that can occur during spec parsing and validation.
 * String values are stable (used in telemetry).
 */
export enum ErrorType {
  SpecNotValid = "spec-not-valid",
  RemoteRefNotSupported = "remote-ref-not-supported",
  NoServerInformation = "no-server-information",
  UrlProtocolNotSupported = "url-protocol-not-supported",
  RelativeServerUrlNotSupported = "relative-server-url-not-supported",
  NoSupportedApi = "no-supported-api",
  NoExtraAPICanBeAdded = "no-extra-api-can-be-added",
  AddedAPINotInOriginalSpec = "added-api-not-in-original-spec",
  ResolveServerUrlFailed = "resolve-server-url-failed",
  SwaggerNotSupported = "swagger-not-supported",
  MultipleAuthNotSupported = "multiple-auth-not-supported",
  SpecVersionNotSupported = "spec-version-not-supported",
  CircularReferenceNotSupported = "circular-reference-not-supported",
  ListFailed = "list-failed",
  FilterSpecFailed = "filter-spec-failed",
  UpdateManifestFailed = "update-manifest-failed",
  GenerateAdaptiveCardFailed = "generate-adaptive-card-failed",
  GenerateFailed = "generate-failed",
  ValidateFailed = "validate-failed",
  GetSpecFailed = "get-spec-failed",
  AuthTypeIsNotSupported = "auth-type-is-not-supported",
  MissingOperationId = "missing-operation-id",
  PostBodyContainMultipleMediaTypes = "post-body-contain-multiple-media-types",
  ResponseContainMultipleMediaTypes = "response-contain-multiple-media-types",
  ResponseJsonIsEmpty = "response-json-is-empty",
  PostBodyContainsRequiredUnsupportedSchema = "post-body-contains-required-unsupported-schema",
  ParamsContainRequiredUnsupportedSchema = "params-contain-required-unsupported-schema",
  ExceededRequiredParamsLimit = "exceeded-required-params-limit",
  NoParameter = "no-parameter",
  NoAPIInfo = "no-api-info",
  MethodNotAllowed = "method-not-allowed",
  UrlPathNotExist = "url-path-not-exist",
  Cancelled = "cancelled",
  Unknown = "unknown",
  AddAuthFailed = "add-auth-failed",
}

export enum WarningType {
  OperationIdMissing = "operationid-missing",
  GenerateCardFailed = "generate-card-failed",
  OperationOnlyContainsOptionalParam = "operation-only-contains-optional-param",
  ConvertSwaggerToOpenAPI = "convert-swagger-to-openapi",
  OperationIdContainsSpecialCharacters = "operationid-contains-special-characters",
  UnsupportedAuthType = "unsupported-auth-type",
  GenerateJsonDataFailed = "generate-json-data-failed",
  Unknown = "unknown",
}

// ---------------------------------------------------------------------------
// Project type
// ---------------------------------------------------------------------------

export enum ProjectType {
  Copilot = "Copilot",
  SME = "SME",
  TeamsAi = "TeamsAi",
}

// ---------------------------------------------------------------------------
// Parse options
// ---------------------------------------------------------------------------

export interface ParseOptions {
  allowMissingId?: boolean;
  allowSwagger?: boolean;
  allowAPIKeyAuth?: boolean;
  allowBearerTokenAuth?: boolean;
  allowMultipleParameters?: boolean;
  allowOauth2?: boolean;
  allowMethods?: string[];
  allowConversationStarters?: boolean;
  allowResponseSemantics?: boolean;
  allowConfirmation?: boolean;
  projectType?: ProjectType;
  isGptPlugin?: boolean;
}

export const DEFAULT_PARSE_OPTIONS: Required<ParseOptions> = {
  allowMissingId: true,
  allowSwagger: true,
  allowAPIKeyAuth: false,
  allowBearerTokenAuth: false,
  allowMultipleParameters: false,
  allowOauth2: false,
  allowMethods: ["get", "post"],
  allowConversationStarters: false,
  allowResponseSemantics: false,
  allowConfirmation: false,
  projectType: ProjectType.SME,
  isGptPlugin: false,
};

// ---------------------------------------------------------------------------
// API listing
// ---------------------------------------------------------------------------

export interface ListAPIInfo {
  api: string;
  server: string;
  operationId: string;
  isValid: boolean;
  reason: ErrorType[];
  auth?: AuthInfo;
  summary?: string;
  description?: string;
}

export interface ListAPIResult {
  allAPICount: number;
  validAPICount: number;
  APIs: ListAPIInfo[];
}

export interface APIValidationResult {
  isValid: boolean;
  reason: ErrorType[];
}

export interface APIMap {
  [key: string]: {
    operation: OpenAPIV3.OperationObject;
    isValid: boolean;
    reason: ErrorType[];
  };
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------

export type AuthType = OpenAPIV3.SecuritySchemeObject | { type: "multipleAuth" };

export interface AuthInfo {
  authScheme: AuthType;
  name: string;
}

export interface OperationAuthInfoMap {
  [operationId: string]: AuthInfo;
}

// ---------------------------------------------------------------------------
// Adaptive Cards
// ---------------------------------------------------------------------------

export interface TextBlockElement {
  type: string;
  text: string;
  wrap: boolean;
}

export interface ImageElement {
  type: string;
  url: string;
  $when: string;
}

export type AdaptiveCardBody = Array<TextBlockElement | ImageElement | ArrayElement>;

export interface ArrayElement {
  type: string;
  $data: string;
  items: AdaptiveCardBody;
}

export interface AdaptiveCard {
  type: string;
  $schema: string;
  version: string;
  body: AdaptiveCardBody;
}

export interface PreviewCardTemplate {
  title: string;
  subtitle?: string;
  image?: {
    url: string;
    alt?: string;
    $when?: string;
  };
}

export interface WrappedAdaptiveCard {
  version: string;
  $schema?: string;
  jsonPath?: string;
  responseLayout: string;
  responseCardTemplate: AdaptiveCard;
  previewCardTemplate: PreviewCardTemplate;
}

// ---------------------------------------------------------------------------
// Generation
// ---------------------------------------------------------------------------

export interface GenerateResult {
  allSuccess: boolean;
  warnings: ValidationWarning[];
}

export interface CheckParamResult {
  requiredNum: number;
  optionalNum: number;
  isValid: boolean;
  reason: ErrorType[];
}

/**
 * Parameter info returned by generateParametersFromSchema.
 */
export interface ParamInfo {
  name: string;
  title: string;
  description: string;
}

export interface InvalidAPIInfo {
  api: string;
  reason: ErrorType[];
}

/** Alias for text block elements */
export type TextElement = TextBlockElement;
