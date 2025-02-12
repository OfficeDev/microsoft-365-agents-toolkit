// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { IParameter } from "@microsoft/teams-manifest";
import { OpenAPIV3 } from "openapi-types";

/**
 * An interface that represents the result of validating an OpenAPI specification file.
 */
export interface ValidateResult {
  /**
   * The validation status of the OpenAPI specification file.
   */
  status: ValidationStatus;

  /**
   * An array of warning results generated during validation.
   */
  warnings: WarningResult[];

  /**
   * An array of error results generated during validation.
   */
  errors: ErrorResult[];

  specHash?: string;
}

export interface SpecValidationResult {
  /**
   * An array of warning results generated during validation.
   */
  warnings: WarningResult[];

  /**
   * An array of error results generated during validation.
   */
  errors: ErrorResult[];
}

/**
 * An interface that represents a warning result generated during validation.
 */
export interface WarningResult {
  /**
   * The type of warning.
   */
  type: WarningType;

  /**
   * The content of the warning.
   */
  content: string;

  /**
   * data of the warning.
   */
  data?: any;
}

/**
 * An interface that represents an error result generated during validation.
 */
export interface ErrorResult {
  /**
   * The type of error.
   */
  type: ErrorType;

  /**
   * The content of the error.
   */
  content: string;

  /**
   * data of the error.
   */
  data?: any;
}

export interface GenerateResult {
  allSuccess: boolean;
  warnings: WarningResult[];
}

/**
 * An enum that represents the types of errors that can occur during validation.
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
  listSupportedAPIInfoFailed = "list-supported-api-info-failed",
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
}

/**
 * An enum that represents the types of warnings that can occur during validation.
 */
export enum WarningType {
  OperationIdMissing = "operationid-missing",
  GenerateCardFailed = "generate-card-failed",
  OperationOnlyContainsOptionalParam = "operation-only-contains-optional-param",
  ConvertSwaggerToOpenAPI = "convert-swagger-to-openapi",
  FuncDescriptionTooLong = "function-description-too-long",
  OperationIdContainsSpecialCharacters = "operationid-contains-special-characters",
  UnsupportedAuthType = "unsupported-auth-type",
  GenerateJsonDataFailed = "generate-json-data-failed",
  Unknown = "unknown",
}

/**
 * An enum that represents the validation status of an OpenAPI specification file.
 */
export enum ValidationStatus {
  Valid,
  Warning, // If there are any warnings, the file is still valid
  Error, // If there are any errors, the file is not valid
}

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

export interface CheckParamResult {
  requiredNum: number;
  optionalNum: number;
  isValid: boolean;
  reason: ErrorType[];
}

export interface ParseOptions {
  /**
   * If true, the parser will not throw an error if an ID is missing the spec file.
   */
  allowMissingId?: boolean;

  /**
   * If true, the parser will allow parsing of Swagger specifications.
   */
  allowSwagger?: boolean;

  /**
   * If true, the parser will allow API Key authentication in the spec file.
   */
  allowAPIKeyAuth?: boolean;

  /**
   * If true, the parser will allow Bearer Token authentication in the spec file.
   */
  allowBearerTokenAuth?: boolean;

  /**
   * If true, the parser will allow multiple parameters in the spec file. Teams AI project would ignore this parameters and always true
   */
  allowMultipleParameters?: boolean;

  /**
   * If true, the parser will allow OAuth2 authentication in the spec file. Currently only support OAuth2 with auth code flow.
   */
  allowOauth2?: boolean;

  /**
   * An array of HTTP methods that the parser will allow in the spec file.
   */
  allowMethods?: string[];

  /**
   * If true, the parser will allow conversation starters in plugin file. Only take effect in Copilot project
   */
  allowConversationStarters?: boolean;

  /**
   * If true, the parser will allow response semantics in plugin file. Only take effect in Copilot project
   */
  allowResponseSemantics?: boolean;

  /**
   * If true, the paser will allow confirmation in plugin file. Only take effect in Copilot project
   */
  allowConfirmation?: boolean;

  /**
   * The type of project that the parser is being used for.
   * Project can be SME/Copilot/TeamsAi
   */
  projectType?: ProjectType;

  /**
   * If true, we will generate files of plugin for GPT (Declarative Extensions in a Copilot Extension). Otherwise, we will generate files of plugin for Copilot.
   */
  isGptPlugin?: boolean;
}

export enum ProjectType {
  Copilot,
  SME,
  TeamsAi,
}

export interface APIInfo {
  method: string;
  path: string;
  title: string;
  id: string;
  parameters: IParameter[];
  description: string;
  warning?: WarningResult;
}

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

export interface APIMap {
  [key: string]: {
    operation: OpenAPIV3.OperationObject;
    isValid: boolean;
    reason: ErrorType[];
  };
}

export interface APIValidationResult {
  isValid: boolean;
  reason: ErrorType[];
}

export interface ListAPIResult {
  allAPICount: number;
  validAPICount: number;
  APIs: ListAPIInfo[];
}

export type AuthType = OpenAPIV3.SecuritySchemeObject | { type: "multipleAuth" };

export interface AuthInfo {
  authScheme: AuthType;
  name: string;
}

export interface InvalidAPIInfo {
  api: string;
  reason: ErrorType[];
}

export interface InferredProperties {
  title?: string;
  subtitle?: string;
  imageUrl?: string;
}

export interface ExistingPluginManifestInfo {
  manifestPath: string;
  specPath: string;
}

export interface OperationAuthInfoMap {
  [operationId: string]: AuthInfo;
}

export interface FunctionClassificationMap {
  [authType: string]: {
    functionNames: string[];
    authName: string;
  };
}
