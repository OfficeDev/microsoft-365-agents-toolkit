// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Constants for spec-parser module.
 * Ported from @microsoft/m365-spec-parser constants.ts.
 */

// ---------------------------------------------------------------------------
// Error / warning messages
// ---------------------------------------------------------------------------

export const SpecParserMessages = {
  CancelledMessage: "Operation is cancelled.",
  NoServerInformation: "No server information is found in the OpenAPI description document.",
  RemoteRefNotSupported: "Remote reference is not supported: %s.",
  MissingOperationId: "Missing operationIds: %s.",
  NoSupportedApi:
    "No supported API is found in the OpenAPI description document: only GET and POST methods are supported, additionally, there can be at most one required parameter, and no auth is allowed.",
  UrlProtocolNotSupported:
    "Server url is not correct: protocol %s is not supported, you should use https protocol instead.",
  RelativeServerUrlNotSupported: "Server url is not correct: relative server url is not supported.",
  ResolveServerUrlFailed:
    "Unable to resolve the server URL: please make sure that the environment variable %s is defined.",
  OperationOnlyContainsOptionalParam:
    "Operation %s contains multiple optional parameters. The first optional parameter is used for this command.",
  ConvertSwaggerToOpenAPI: "The Swagger 2.0 file has been converted to OpenAPI 3.0.",
  SwaggerNotSupported:
    "Swagger 2.0 is not supported. Please convert to OpenAPI 3.0 manually before proceeding.",
  SpecVersionNotSupported: "Unsupported OpenAPI version %s. Please use version 3.0.x.",
  MultipleAuthNotSupported:
    "Multiple authentication methods are unsupported. Ensure all selected APIs use identical authentication.",
  OperationIdContainsSpecialCharacters:
    "Operation id '%s' in OpenAPI description document contained special characters and was renamed to '%s'.",
  AuthTypeIsNotSupported:
    "Unsupported authorization type in API '%s'. No authorization will be used.",
  UnsupportedSchema: "Unsupported schema in %s %s: %s",
  FuncDescriptionTooLong:
    "The description of the function '%s' is too long. The current length is %s characters, while the maximum allowed length is %s characters.",
  GenerateJsonDataFailed: "Failed to generate JSON data for api: %s due to %s.",
  SchemaNotSupported: "'oneOf', 'allOf', 'anyOf', and 'not' schema are not supported: %s.",
  UnknownSchema: "Unknown schema: %s.",
  AdditionalPropertiesNotSupported: "'additionalProperties' is not supported, and will be ignored.",
} as const;

// ---------------------------------------------------------------------------
// Adaptive Card constants
// ---------------------------------------------------------------------------

export const AdaptiveCardConstants = {
  WrappedCardVersion: "1.0",
  WrappedCardSchema:
    "https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.ResponseRenderingTemplate.schema.json",
  WrappedCardResponseLayout: "list",
  AdaptiveCardVersion: "1.5",
  AdaptiveCardSchema: "https://adaptivecards.io/schemas/adaptive-card.json",
  AdaptiveCardType: "AdaptiveCard",
  TextBlockType: "TextBlock",
  ImageType: "Image",
  ContainerType: "Container",
} as const;

// ---------------------------------------------------------------------------
// Protocol & methods
// ---------------------------------------------------------------------------

export const HTTPMethods = {
  Get: "get",
  Post: "post",
  AllOperationMethods: ["get", "post", "put", "delete", "patch", "head", "options", "trace"],
} as const;

// ---------------------------------------------------------------------------
// Response codes
// ---------------------------------------------------------------------------

export const ResponseCodeFor20X = [
  "200",
  "201",
  "202",
  "203",
  "204",
  "205",
  "206",
  "207",
  "208",
  "226",
  "2XX",
  "default",
] as const;

// ---------------------------------------------------------------------------
// Well-known property names (for Adaptive Card inference)
// ---------------------------------------------------------------------------

export const WellKnownNames = {
  Results: ["result", "data", "items", "root", "matches", "queries", "list", "output"],
  Title: ["title", "name", "summary", "caption", "subject", "label"],
  Subtitle: ["subtitle", "id", "uid", "description", "desc", "detail"],
  Image: ["image", "icon", "avatar", "picture", "photo", "logo", "pic", "thumbnail", "img"],
} as const;

// ---------------------------------------------------------------------------
// Length limits
// ---------------------------------------------------------------------------

export const Limits = {
  ShortDescriptionMaxLen: 80,
  FullDescriptionMaxLen: 4000,
  CommandDescriptionMaxLen: 128,
  ParameterDescriptionMaxLen: 128,
  ConversationStarterMaxLen: 50,
  CommandTitleMaxLen: 32,
  ParameterTitleMaxLen: 32,
  SMERequiredParamsMaxNum: 5,
  FunctionDescriptionMaxLen: 100,
} as const;

export const RegistrationIdPostfix = "REGISTRATION_ID";
export const DefaultPluginId = "plugin_1";
