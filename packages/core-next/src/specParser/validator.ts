// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * OpenAPI spec validators — per-project-type validation of operations and specs.
 *
 * Each project type (Copilot, SME, TeamsAi) enforces different rules about
 * auth, parameters, responses, circular references, etc.
 */

import type { OpenAPIV3 } from "openapi-types";
import type {
  ParseOptions,
  APIValidationResult,
  ValidationError,
  ValidationWarning,
  APIMap,
  InvalidAPIInfo,
} from "./types";
import { ErrorType, WarningType, ProjectType } from "./types";
import { SpecParserMessages } from "./constants";
import {
  getAuthArray,
  isAPIKeyAuth,
  isOAuthWithAuthCodeFlow,
  isBearerTokenAuth,
  containMultipleMediaTypes,
  getResponseJson,
  isObjectSchema,
  getServerObject,
  checkServerUrl,
  validateServer,
  formatStr,
} from "./utils";

// ---------------------------------------------------------------------------
// SpecValidationResult — shared shape for spec-level validation
// ---------------------------------------------------------------------------

export interface SpecValidationResult {
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

// ---------------------------------------------------------------------------
// Abstract base
// ---------------------------------------------------------------------------

export abstract class Validator {
  readonly projectType: ProjectType;
  readonly spec: OpenAPIV3.Document;
  readonly options: ParseOptions;

  private _apiMap: APIMap | undefined;
  protected hasCircularReference = false;

  constructor(spec: OpenAPIV3.Document, options: ParseOptions, projectType: ProjectType) {
    this.spec = spec;
    this.options = options;
    this.projectType = projectType;
  }

  abstract validateAPI(method: string, path: string): APIValidationResult;
  abstract validateSpec(): SpecValidationResult;

  // -----------------------------------------------------------------------
  // Shared helpers
  // -----------------------------------------------------------------------

  listAPIs(): APIMap {
    if (this._apiMap) return this._apiMap;

    const result: APIMap = {};
    for (const path in this.spec.paths) {
      const methods = this.spec.paths[path]!;
      for (const method in methods) {
        const op = (methods as Record<string, unknown>)[method] as OpenAPIV3.OperationObject;
        if (this.options.allowMethods?.includes(method) && op) {
          const v = this.validateAPI(method, path);
          result[`${method.toUpperCase()} ${path}`] = {
            operation: op,
            isValid: v.isValid,
            reason: v.reason,
          };
        }
      }
    }
    this._apiMap = result;
    return result;
  }

  protected checkCircularReference(): void {
    try {
      JSON.stringify(this.spec);
    } catch (e) {
      if ((e as Error).message.includes("Converting circular structure to JSON")) {
        this.hasCircularReference = true;
      }
    }
  }

  protected validateSpecVersion(): SpecValidationResult {
    const r: SpecValidationResult = { errors: [], warnings: [] };
    if (this.spec.openapi >= "3.1.0") {
      r.errors.push({
        type: ErrorType.SpecVersionNotSupported,
        content: formatStr(SpecParserMessages.SpecVersionNotSupported, this.spec.openapi),
        data: this.spec.openapi,
      });
    }
    return r;
  }

  protected validateSpecServer(): SpecValidationResult {
    return { errors: validateServer(this.spec, this.options), warnings: [] };
  }

  protected validateSpecNoSupportAPI(): SpecValidationResult {
    const r: SpecValidationResult = { errors: [], warnings: [] };
    const apiMap = this.listAPIs();
    const validAPIs = Object.entries(apiMap).filter(([, v]) => v.isValid);
    if (validAPIs.length === 0) {
      const data: InvalidAPIInfo[] = Object.entries(apiMap).map(([api, { reason }]) => ({
        api,
        reason,
      }));
      r.errors.push({
        type: ErrorType.NoSupportedApi,
        content: SpecParserMessages.NoSupportedApi,
        data,
      });
    }
    return r;
  }

  protected validateSpecOperationId(): SpecValidationResult {
    const r: SpecValidationResult = { errors: [], warnings: [] };
    const apiMap = this.listAPIs();
    const missing = Object.entries(apiMap)
      .filter(([, v]) => !v.operation.operationId)
      .map(([key]) => key);
    if (missing.length > 0) {
      r.warnings.push({
        type: WarningType.OperationIdMissing,
        content: formatStr(SpecParserMessages.MissingOperationId, missing.join(", ")),
        data: missing,
      });
    }
    return r;
  }

  protected validateMethodAndPath(method: string, path: string): APIValidationResult {
    const r: APIValidationResult = { isValid: true, reason: [] };
    if (this.options.allowMethods && !this.options.allowMethods.includes(method)) {
      return { isValid: false, reason: [ErrorType.MethodNotAllowed] };
    }
    const pathObj = this.spec.paths[path] as Record<string, unknown> | undefined;
    if (!pathObj || !pathObj[method]) {
      return { isValid: false, reason: [ErrorType.UrlPathNotExist] };
    }
    return r;
  }

  protected validateCircularRef(method: string, path: string): APIValidationResult {
    const r: APIValidationResult = { isValid: true, reason: [] };
    if (this.hasCircularReference) {
      const op = (this.spec.paths[path] as Record<string, unknown>)[
        method
      ] as OpenAPIV3.OperationObject;
      try {
        JSON.stringify(op);
      } catch (e) {
        if ((e as Error).message.includes("Converting circular structure to JSON")) {
          r.isValid = false;
          r.reason.push(ErrorType.CircularReferenceNotSupported);
        }
      }
    }
    return r;
  }

  protected validateServerForAPI(method: string, path: string): APIValidationResult {
    const r: APIValidationResult = { isValid: true, reason: [] };
    const serverObj = getServerObject(this.spec, method, path);
    if (!serverObj) {
      r.reason.push(ErrorType.NoServerInformation);
    } else {
      const allowHttp = this.projectType === ProjectType.Copilot;
      r.reason.push(...checkServerUrl([serverObj], allowHttp).map((e) => e.type));
    }
    return r;
  }

  protected validateAuth(method: string, path: string): APIValidationResult {
    const op = (this.spec.paths[path] as Record<string, unknown>)[
      method
    ] as OpenAPIV3.OperationObject;
    const authSchemeArray = getAuthArray(op.security, this.spec);

    if (authSchemeArray.length === 0) return { isValid: true, reason: [] };

    if (
      this.options.allowAPIKeyAuth ||
      this.options.allowOauth2 ||
      this.options.allowBearerTokenAuth
    ) {
      if (authSchemeArray.every((auths) => auths.length > 1)) {
        return { isValid: false, reason: [ErrorType.MultipleAuthNotSupported] };
      }

      if (this.projectType === ProjectType.Copilot) {
        return { isValid: true, reason: [] };
      }

      for (const auths of authSchemeArray) {
        if (auths.length === 1) {
          const scheme = auths[0].authScheme;
          if (
            (this.options.allowAPIKeyAuth && isAPIKeyAuth(scheme)) ||
            (this.options.allowOauth2 && isOAuthWithAuthCodeFlow(scheme)) ||
            (this.options.allowBearerTokenAuth && isBearerTokenAuth(scheme))
          ) {
            return { isValid: true, reason: [] };
          }
        }
      }
    }

    return { isValid: false, reason: [ErrorType.AuthTypeIsNotSupported] };
  }
}

// ---------------------------------------------------------------------------
// CopilotValidator
// ---------------------------------------------------------------------------

export class CopilotValidator extends Validator {
  constructor(spec: OpenAPIV3.Document, options: ParseOptions) {
    super(spec, options, ProjectType.Copilot);
  }

  validateSpec(): SpecValidationResult {
    const r: SpecValidationResult = { errors: [], warnings: [] };
    r.errors.push(...this.validateSpecVersion().errors);
    r.errors.push(...this.validateSpecServer().errors);
    r.errors.push(...this.validateSpecNoSupportAPI().errors);
    r.warnings.push(...this.validateSpecOperationId().warnings);
    return r;
  }

  validateAPI(method: string, path: string): APIValidationResult {
    const r: APIValidationResult = { isValid: true, reason: [] };
    method = method.toLowerCase();

    const mp = this.validateMethodAndPath(method, path);
    if (!mp.isValid) return mp;

    const op = (this.spec.paths[path] as Record<string, unknown>)[
      method
    ] as OpenAPIV3.OperationObject;

    r.reason.push(...this.validateAuth(method, path).reason);
    if (!this.options.allowMissingId && !op.operationId)
      r.reason.push(ErrorType.MissingOperationId);
    r.reason.push(...this.validateServerForAPI(method, path).reason);

    if (r.reason.length > 0) r.isValid = false;
    return r;
  }
}

// ---------------------------------------------------------------------------
// SMEValidator
// ---------------------------------------------------------------------------

interface ParamCheckResult {
  requiredNum: number;
  optionalNum: number;
  isValid: boolean;
  reason: ErrorType[];
}

const SME_REQUIRED_PARAMS_MAX = 5;

export class SMEValidator extends Validator {
  constructor(spec: OpenAPIV3.Document, options: ParseOptions) {
    super(spec, options, ProjectType.SME);
    this.checkCircularReference();
  }

  validateSpec(): SpecValidationResult {
    const r: SpecValidationResult = { errors: [], warnings: [] };
    r.errors.push(...this.validateSpecVersion().errors);
    r.errors.push(...this.validateSpecServer().errors);
    r.errors.push(...this.validateSpecNoSupportAPI().errors);
    if (this.options.allowMissingId) {
      r.warnings.push(...this.validateSpecOperationId().warnings);
    }
    return r;
  }

  validateAPI(method: string, path: string): APIValidationResult {
    const r: APIValidationResult = { isValid: true, reason: [] };
    method = method.toLowerCase();

    const mp = this.validateMethodAndPath(method, path);
    if (!mp.isValid) return mp;

    const cr = this.validateCircularRef(method, path);
    if (!cr.isValid) return cr;

    const op = (this.spec.paths[path] as Record<string, unknown>)[
      method
    ] as OpenAPIV3.OperationObject;

    r.reason.push(...this.validateAuth(method, path).reason);
    if (!this.options.allowMissingId && !op.operationId)
      r.reason.push(ErrorType.MissingOperationId);
    r.reason.push(...this.validateServerForAPI(method, path).reason);

    // Response validation
    const { json, multipleMediaType } = getResponseJson(op);
    if (multipleMediaType) {
      r.reason.push(ErrorType.ResponseContainMultipleMediaTypes);
    } else if (Object.keys(json).length === 0) {
      r.reason.push(ErrorType.ResponseJsonIsEmpty);
    }

    // Request body
    let postBodyResult: ParamCheckResult = {
      requiredNum: 0,
      optionalNum: 0,
      isValid: true,
      reason: [],
    };
    const requestBody = op.requestBody as OpenAPIV3.RequestBodyObject | undefined;
    if (requestBody) {
      if (containMultipleMediaTypes(requestBody)) {
        r.reason.push(ErrorType.PostBodyContainMultipleMediaTypes);
      }
      const requestJson = requestBody.content?.["application/json"];
      if (requestJson) {
        postBodyResult = this.checkPostBodySchema(
          requestJson.schema as OpenAPIV3.SchemaObject,
          requestBody.required
        );
        r.reason.push(...postBodyResult.reason);
      }
    }

    // Parameters
    const paramObject = op.parameters as OpenAPIV3.ParameterObject[] | undefined;
    const paramResult = this.checkParamSchema(paramObject);
    r.reason.push(...paramResult.reason);

    // Total param count check
    if (paramResult.isValid && postBodyResult.isValid) {
      const totalRequired = postBodyResult.requiredNum + paramResult.requiredNum;
      const totalParams = totalRequired + postBodyResult.optionalNum + paramResult.optionalNum;

      if (totalRequired > 1) {
        if (!this.options.allowMultipleParameters || totalRequired > SME_REQUIRED_PARAMS_MAX) {
          r.reason.push(ErrorType.ExceededRequiredParamsLimit);
        }
      } else if (totalParams === 0) {
        r.reason.push(ErrorType.NoParameter);
      }
    }

    if (r.reason.length > 0) r.isValid = false;
    return r;
  }

  private checkPostBodySchema(
    schema: OpenAPIV3.SchemaObject,
    isRequired = false
  ): ParamCheckResult {
    const result: ParamCheckResult = { requiredNum: 0, optionalNum: 0, isValid: true, reason: [] };
    if (Object.keys(schema).length === 0) return result;

    const isRequiredNoDefault = isRequired && schema.default === undefined;

    if (["string", "integer", "boolean", "number"].includes(schema.type as string)) {
      if (isRequiredNoDefault) result.requiredNum++;
      else result.optionalNum++;
    } else if (isObjectSchema(schema)) {
      for (const prop in schema.properties) {
        const propRequired = schema.required?.includes(prop) ?? false;
        const sub = this.checkPostBodySchema(
          schema.properties[prop] as OpenAPIV3.SchemaObject,
          propRequired
        );
        result.requiredNum += sub.requiredNum;
        result.optionalNum += sub.optionalNum;
        result.isValid = result.isValid && sub.isValid;
        result.reason.push(...sub.reason);
      }
    } else if (isRequiredNoDefault) {
      result.isValid = false;
      result.reason.push(ErrorType.PostBodyContainsRequiredUnsupportedSchema);
    }
    return result;
  }

  private checkParamSchema(paramObject?: OpenAPIV3.ParameterObject[]): ParamCheckResult {
    const result: ParamCheckResult = { requiredNum: 0, optionalNum: 0, isValid: true, reason: [] };
    if (!paramObject) return result;

    for (const param of paramObject) {
      const schema = param.schema as OpenAPIV3.SchemaObject;
      const isRequiredNoDefault = param.required && schema.default === undefined;

      if (param.in === "header" || param.in === "cookie") {
        if (isRequiredNoDefault) {
          result.isValid = false;
          result.reason.push(ErrorType.ParamsContainRequiredUnsupportedSchema);
        }
        continue;
      }

      if (!["boolean", "string", "number", "integer"].includes(schema.type as string)) {
        if (isRequiredNoDefault) {
          result.isValid = false;
          result.reason.push(ErrorType.ParamsContainRequiredUnsupportedSchema);
        }
        continue;
      }

      if (param.in === "query" || param.in === "path") {
        if (isRequiredNoDefault) result.requiredNum++;
        else result.optionalNum++;
      }
    }
    return result;
  }
}

// ---------------------------------------------------------------------------
// TeamsAIValidator
// ---------------------------------------------------------------------------

export class TeamsAIValidator extends Validator {
  constructor(spec: OpenAPIV3.Document, options: ParseOptions) {
    super(spec, options, ProjectType.TeamsAi);
    this.checkCircularReference();
  }

  validateSpec(): SpecValidationResult {
    const r: SpecValidationResult = { errors: [], warnings: [] };
    r.errors.push(...this.validateSpecServer().errors);
    r.errors.push(...this.validateSpecNoSupportAPI().errors);
    return r;
  }

  validateAPI(method: string, path: string): APIValidationResult {
    const r: APIValidationResult = { isValid: true, reason: [] };
    method = method.toLowerCase();

    const mp = this.validateMethodAndPath(method, path);
    if (!mp.isValid) return mp;

    const cr = this.validateCircularRef(method, path);
    if (!cr.isValid) return cr;

    const op = (this.spec.paths[path] as Record<string, unknown>)[
      method
    ] as OpenAPIV3.OperationObject;

    if (!this.options.allowMissingId && !op.operationId)
      r.reason.push(ErrorType.MissingOperationId);
    r.reason.push(...this.validateServerForAPI(method, path).reason);

    if (r.reason.length > 0) r.isValid = false;
    return r;
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

export function createValidator(spec: OpenAPIV3.Document, options: ParseOptions): Validator {
  const type = options.projectType ?? ProjectType.SME;
  switch (type) {
    case ProjectType.Copilot:
      return new CopilotValidator(spec, options);
    case ProjectType.TeamsAi:
      return new TeamsAIValidator(spec, options);
    case ProjectType.SME:
    default:
      return new SMEValidator(spec, options);
  }
}
