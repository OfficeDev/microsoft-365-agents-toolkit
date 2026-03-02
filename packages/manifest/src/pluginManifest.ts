// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 * Import `APIPluginManifest` or version-specific types like `APIPluginManifestV2D4` from the generated-types module.
 */

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export type Instruction = string | string[];

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export type Example = string | string[];

/**
 * @deprecated Use `APIPluginManifest` or version-specific types from `./generated-types` instead.
 */
export interface PluginManifestSchema {
  schema_version: string;
  name_for_human: string;
  namespace?: string;
  description_for_model?: string;
  description_for_human: string;
  logo_url?: string;
  contact_email?: string;
  legal_info_url?: string;
  privacy_policy_url?: string;
  functions?: FunctionObject[];
  runtimes?: (RuntimeObjectLocalplugin | RuntimeObjectOpenapi)[];
  capabilities?: {
    conversation_starters?: ConversationStarter[];
    [k: string]: unknown;
  };
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface FunctionObject {
  name: string;
  description?: string;
  parameters?: FunctionParameters;
  returns?: FunctionReturnType | FunctionRichResponseReturnType;
  states?: {
    reasoning?: FunctionStateConfig;
    responding?: FunctionStateConfig;
    [k: string]: unknown;
  };
  capabilities?: {
    confirmation?: ConfirmationObject;
    response_semantics?: ResponseSemanticsObject;
    [k: string]: unknown;
  };
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface FunctionParameters {
  type?: "object";
  properties: {
    [k: string]: FunctionParameter;
  };
  required?: string[];
  [k: string]: unknown;
}
/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 * This interface was referenced by `undefined`'s JSON-Schema definition
 * via the `patternProperty` "^[A-Za-z0-9_]+$".
 */
export interface FunctionParameter {
  type: "string" | "array" | "boolean" | "integer" | "number";
  items?: {
    [k: string]: unknown;
  };
  enum?: string[];
  description?: string;
  default?: string | boolean | number | number | unknown[];
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface FunctionReturnType {
  type: "string";
  description?: string;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface FunctionRichResponseReturnType {
  $ref: "https://copilot.microsoft.com/schemas/rich-response-v1.0.json";
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface FunctionStateConfig {
  description?: string;
  instructions?: Instruction;
  examples?: Example;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface ConfirmationObject {
  type?: "None" | "AdaptiveCard";
  title?: string;
  body?: string;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface ResponseSemanticsObject {
  data_path: string;
  properties?: {
    title?: string;
    subtitle?: string;
    url?: string;
    information_protection_label?: string;
    template_selector?: string;
    [k: string]: unknown;
  };
  static_template?: {
    [k: string]: unknown;
  };
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface RuntimeObjectLocalplugin {
  type: "LocalPlugin";
  run_for_functions?: string[];
  spec: LocalPluginRuntime;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface LocalPluginRuntime {
  local_endpoint: string;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface RuntimeObjectOpenapi {
  type: "OpenApi";
  auth?: AuthObject;
  run_for_functions?: string[];
  spec: OpenApiRuntime;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface AuthObject {
  type: "None" | "OAuthPluginVault" | "ApiKeyPluginVault";
  reference_id?: string;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface OpenApiRuntime {
  url: string;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface LocalizationObject {
  /**
   * This interface was referenced by `LocalizationObject`'s JSON-Schema definition
   * via the `patternProperty` "^(?i)[a-z]{2,3}(-[a-z]{2})?(?-i)$".
   */
  [k: string]: {
    /**
     * This interface was referenced by `undefined`'s JSON-Schema definition
     * via the `patternProperty` "^[A-Za-z_][A-Za-z0-9_]*$".
     */
    [k: string]: {
      message: string;
      description: string;
      [k: string]: unknown;
    };
  };
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface ConversationStarter {
  text: string;
  title?: string;
  [k: string]: unknown;
}
