// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Pure utility functions for OpenAPI spec analysis.
 * Ported from @microsoft/m365-spec-parser Utils class as standalone functions.
 */

import type { OpenAPIV3 } from "openapi-types";
import type {
  AuthInfo,
  AuthType,
  OperationAuthInfoMap,
  ValidationError,
  ParseOptions,
  ParamInfo,
} from "./types";
import { ErrorType, ProjectType } from "./types";
import { SpecParserMessages, ResponseCodeFor20X, Limits } from "./constants";

// ---------------------------------------------------------------------------
// Auth helpers
// ---------------------------------------------------------------------------

export function isBearerTokenAuth(authScheme: AuthType): boolean {
  return authScheme.type === "http" && authScheme.scheme?.toLowerCase() === "bearer";
}

export function isAPIKeyAuth(authScheme: AuthType): boolean {
  return authScheme.type === "apiKey";
}

export function isAPIKeyAuthButNotInCookie(authScheme: AuthType): boolean {
  return authScheme.type === "apiKey" && authScheme.in !== "cookie";
}

export function isOAuthWithAuthCodeFlow(authScheme: AuthType): boolean {
  return !!(authScheme.type === "oauth2" && authScheme.flows && authScheme.flows.authorizationCode);
}

export function isNotSupportedAuth(authSchemeArray: AuthInfo[][]): boolean {
  if (authSchemeArray.length === 0) return false;
  if (authSchemeArray.every((auths) => auths.length > 1)) return true;

  for (const auths of authSchemeArray) {
    if (auths.length === 1) {
      if (
        isOAuthWithAuthCodeFlow(auths[0].authScheme) ||
        isBearerTokenAuth(auths[0].authScheme) ||
        isAPIKeyAuthButNotInCookie(auths[0].authScheme)
      ) {
        return false;
      }
    }
  }
  return true;
}

/**
 * Extract auth info arrays from security requirements.
 */
export function getAuthArray(
  securities: OpenAPIV3.SecurityRequirementObject[] | undefined,
  spec: OpenAPIV3.Document
): AuthInfo[][] {
  const result: AuthInfo[][] = [];
  const securitySchemas = spec.components?.securitySchemes;
  const securitiesArr = securities ?? spec.security;

  if (securitiesArr && securitySchemas) {
    for (const security of securitiesArr) {
      const authArray: AuthInfo[] = [];
      for (const name in security) {
        const auth = securitySchemas[name] as OpenAPIV3.SecuritySchemeObject;
        authArray.push({ authScheme: auth, name });
      }
      if (authArray.length > 0) {
        result.push(authArray);
      }
    }
  }

  result.sort((a, b) => a[0].name.localeCompare(b[0].name));
  return result;
}

/**
 * Build operation → auth mapping for the entire spec.
 */
export function getAuthMap(spec: OpenAPIV3.Document): OperationAuthInfoMap {
  const authMap: OperationAuthInfoMap = {};

  for (const url in spec.paths) {
    for (const method in spec.paths[url]) {
      const operation = (spec.paths[url] as Record<string, unknown>)[
        method
      ] as OpenAPIV3.OperationObject;
      const authArray = getAuthArray(operation.security, spec);
      if (authArray.length > 0) {
        authMap[operation.operationId!] = authArray[0][0];
      }
    }
  }
  return authMap;
}

/**
 * Get the single AuthInfo for a spec (errors if multiple different auth found).
 */
export function getAuthInfo(spec: OpenAPIV3.Document): {
  authInfo: AuthInfo | undefined;
  error?: string;
} {
  let authInfo: AuthInfo | undefined;

  for (const url in spec.paths) {
    for (const method in spec.paths[url]) {
      const operation = (spec.paths[url] as Record<string, unknown>)[
        method
      ] as OpenAPIV3.OperationObject;
      const authArray = getAuthArray(operation.security, spec);
      if (authArray.length > 0) {
        const currentAuth = authArray[0][0];
        if (!authInfo) {
          authInfo = currentAuth;
        } else if (authInfo.name !== currentAuth.name) {
          return { authInfo: undefined, error: SpecParserMessages.MultipleAuthNotSupported };
        }
      }
    }
  }

  return { authInfo };
}

// ---------------------------------------------------------------------------
// Schema helpers
// ---------------------------------------------------------------------------

export function isObjectSchema(schema: OpenAPIV3.SchemaObject): boolean {
  return schema.type === "object" || (!schema.type && !!schema.properties);
}

export function containMultipleMediaTypes(
  bodyObject: OpenAPIV3.RequestBodyObject | OpenAPIV3.ResponseObject
): boolean {
  return Object.keys(bodyObject?.content || {}).length > 1;
}

/**
 * Get the JSON response media type from a 2xx response.
 */
export function getResponseJson(
  operationObject: OpenAPIV3.OperationObject | undefined,
  allowMultipleMediaType = false
): { json: OpenAPIV3.MediaTypeObject; multipleMediaType: boolean } {
  let json: OpenAPIV3.MediaTypeObject = {};
  let multipleMediaType = false;

  for (const code of ResponseCodeFor20X) {
    const responseObject = operationObject?.responses?.[code] as OpenAPIV3.ResponseObject;
    if (!responseObject) continue;

    multipleMediaType = containMultipleMediaTypes(responseObject);
    if (!allowMultipleMediaType && multipleMediaType) {
      json = {};
      continue;
    }

    const mediaObj = getJsonContentType(responseObject);
    if (Object.keys(mediaObj).length > 0) {
      json = mediaObj;
      return { json, multipleMediaType };
    }
  }

  return { json, multipleMediaType };
}

export function getJsonContentType(
  responseObject: OpenAPIV3.ResponseObject | OpenAPIV3.RequestBodyObject
): OpenAPIV3.MediaTypeObject {
  if (responseObject.content) {
    for (const contentType of Object.keys(responseObject.content)) {
      if (contentType.indexOf("application/json") >= 0) {
        return responseObject.content[contentType];
      }
    }
  }
  return {};
}

// ---------------------------------------------------------------------------
// Path & naming helpers
// ---------------------------------------------------------------------------

export function convertPathToCamelCase(path: string): string {
  const segments = path.split(/[./{]/);
  return segments
    .map((segment) => {
      segment = segment.replace(/}/g, "");
      return segment.charAt(0).toUpperCase() + segment.slice(1);
    })
    .join("");
}

export function updateFirstLetter(str: string): string {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

/**
 * Convert an auth name to a safe environment variable name.
 */
export function getSafeRegistrationIdEnvName(authName: string): string {
  if (!authName) return "";
  let safe = authName.toUpperCase().replace(/[^A-Z0-9_]/g, "_");
  if (!safe.match(/^[A-Z]/)) {
    safe = "PREFIX_" + safe;
  }
  return safe;
}

// ---------------------------------------------------------------------------
// URL / Server helpers
// ---------------------------------------------------------------------------

export function getUrlProtocol(urlString: string): string | undefined {
  try {
    return new URL(urlString).protocol;
  } catch {
    return undefined;
  }
}

export function resolveEnv(str: string): string {
  const placeHolderReg = /\$\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}/g;
  let matches = placeHolderReg.exec(str);
  let newStr = str;
  while (matches != null) {
    const envVar = matches[1];
    const envVal = process.env[envVar];
    if (!envVal) {
      throw new Error(formatStr(SpecParserMessages.ResolveServerUrlFailed, envVar));
    } else {
      newStr = newStr.replace(matches[0], envVal);
    }
    matches = placeHolderReg.exec(str);
  }
  return newStr;
}

export function checkServerUrl(
  servers: OpenAPIV3.ServerObject[],
  allowHttp = false
): ValidationError[] {
  const errors: ValidationError[] = [];

  let serverUrl: string;
  try {
    serverUrl = resolveEnv(servers[0].url);
  } catch (e) {
    errors.push({
      type: ErrorType.ResolveServerUrlFailed,
      content: (e as Error).message,
      data: servers,
    });
    return errors;
  }

  const protocol = getUrlProtocol(serverUrl);
  if (!protocol) {
    errors.push({
      type: ErrorType.RelativeServerUrlNotSupported,
      content: SpecParserMessages.RelativeServerUrlNotSupported,
      data: servers,
    });
  } else if (protocol !== "https:" && !(protocol === "http:" && allowHttp)) {
    const protocolString = protocol.slice(0, -1);
    errors.push({
      type: ErrorType.UrlProtocolNotSupported,
      content: formatStr(SpecParserMessages.UrlProtocolNotSupported, protocolString),
      data: protocolString,
    });
  }

  return errors;
}

/**
 * Validate server entries across spec, paths, and operations.
 */
export function validateServer(spec: OpenAPIV3.Document, options: ParseOptions): ValidationError[] {
  const errors: ValidationError[] = [];
  const allowHttp = options.projectType === ProjectType.Copilot;

  let hasTopLevelServers = false;
  let hasPathLevelServers = false;
  let hasOperationLevelServers = false;

  if (spec.servers && spec.servers.length >= 1) {
    hasTopLevelServers = true;
    errors.push(...checkServerUrl(spec.servers, allowHttp));
  }

  for (const path in spec.paths) {
    const methods = spec.paths[path]!;
    if (methods.servers && methods.servers.length >= 1) {
      hasPathLevelServers = true;
      errors.push(...checkServerUrl(methods.servers, allowHttp));
    }

    for (const method in methods) {
      const operation = (methods as Record<string, unknown>)[method] as OpenAPIV3.OperationObject;
      if (options.allowMethods?.includes(method) && operation?.servers?.length) {
        hasOperationLevelServers = true;
        errors.push(...checkServerUrl(operation.servers, allowHttp));
      }
    }
  }

  if (!hasTopLevelServers && !hasPathLevelServers && !hasOperationLevelServers) {
    errors.push({
      type: ErrorType.NoServerInformation,
      content: SpecParserMessages.NoServerInformation,
    });
  }

  return errors;
}

/**
 * Get the resolved server object for a specific method/path.
 * Precedence: operation > path > root.
 */
export function getServerObject(
  spec: OpenAPIV3.Document,
  method: string,
  path: string
): OpenAPIV3.ServerObject | undefined {
  const pathObj = spec.paths[path] as Record<string, unknown>;
  const operation = pathObj[method] as OpenAPIV3.OperationObject;

  return (
    (operation.servers && operation.servers[0]) ||
    (spec.paths[path]!.servers && spec.paths[path]!.servers![0]) ||
    (spec.servers && spec.servers[0])
  );
}

// ---------------------------------------------------------------------------
// Parameter helpers
// ---------------------------------------------------------------------------

export function generateParametersFromSchema(
  schema: OpenAPIV3.SchemaObject,
  name: string,
  allowMultipleParameters: boolean,
  isRequired = false
): { requiredParams: ParamInfo[]; optionalParams: ParamInfo[] } {
  const requiredParams: ParamInfo[] = [];
  const optionalParams: ParamInfo[] = [];

  if (["string", "integer", "boolean", "number"].includes(schema.type as string)) {
    const parameter: ParamInfo = {
      name,
      title: updateFirstLetter(name).slice(0, Limits.ParameterTitleMaxLen),
      description: (schema.description ?? "").slice(0, Limits.ParameterDescriptionMaxLen),
    };

    if (isRequired && schema.default === undefined) {
      requiredParams.push(parameter);
    } else {
      optionalParams.push(parameter);
    }
  } else if (isObjectSchema(schema)) {
    for (const property in schema.properties) {
      const propRequired = schema.required?.includes(property) ?? false;
      const result = generateParametersFromSchema(
        schema.properties[property] as OpenAPIV3.SchemaObject,
        property,
        allowMultipleParameters,
        propRequired
      );
      requiredParams.push(...result.requiredParams);
      optionalParams.push(...result.optionalParams);
    }
  }

  return { requiredParams, optionalParams };
}

// ---------------------------------------------------------------------------
// Well-known name detection
// ---------------------------------------------------------------------------

export function isWellKnownName(name: string, wellknownNameList: string[]): boolean {
  const normalized = name.replace(/_/g, "").replace(/-/g, "").toLowerCase();
  return wellknownNameList.some((wn) => normalized.includes(wn));
}

// ---------------------------------------------------------------------------
// String formatting
// ---------------------------------------------------------------------------

export function formatStr(str: string, ...args: string[]): string {
  let index = 0;
  return str.replace(/%s/g, () => args[index++] ?? "");
}

// ---------------------------------------------------------------------------
// Security scheme construction
// ---------------------------------------------------------------------------

export function getAuthSchemaObject(
  authType: string,
  authParameters: Record<string, unknown>
): OpenAPIV3.SecuritySchemeObject {
  switch (authType) {
    case "oauth":
    case "microsoft-entra":
      return {
        type: "oauth2",
        flows: {
          authorizationCode: {
            authorizationUrl: authParameters.authorizationUrl as string,
            tokenUrl: authParameters.tokenUrl as string,
            refreshUrl: authParameters.refreshUrl as string,
            scopes: authParameters.scopes as Record<string, string>,
          },
        },
      } as OpenAPIV3.OAuth2SecurityScheme;
    case "api-key":
      return {
        type: "apiKey",
        in: authParameters.in as string,
        name: authParameters.name as string,
      } as OpenAPIV3.ApiKeySecurityScheme;
    case "bearer-token":
    default:
      return {
        type: "http",
        scheme: "bearer",
      } as OpenAPIV3.HttpSecurityScheme;
  }
}
