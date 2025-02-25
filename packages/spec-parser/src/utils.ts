// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { OpenAPIV3 } from "openapi-types";
import { ConstantString } from "./constants";
import {
  AdaptiveCardBody,
  ArrayElement,
  AuthInfo,
  AuthType,
  ErrorResult,
  ErrorType,
  OperationAuthInfoMap,
  ParseOptions,
  ProjectType,
} from "./interfaces";
import { IMessagingExtensionCommand, IParameter } from "@microsoft/teams-manifest";
import { SpecParserError } from "./specParserError";

export class Utils {
  static isObjectSchema(schema: OpenAPIV3.SchemaObject): boolean {
    return schema.type === "object" || (!schema.type && !!schema.properties);
  }

  static containMultipleMediaTypes(
    bodyObject: OpenAPIV3.RequestBodyObject | OpenAPIV3.ResponseObject
  ): boolean {
    return Object.keys(bodyObject?.content || {}).length > 1;
  }

  static isBearerTokenAuth(authScheme: AuthType): boolean {
    return authScheme.type === "http" && authScheme.scheme === "bearer";
  }

  static isAPIKeyAuth(authScheme: AuthType): boolean {
    return authScheme.type === "apiKey";
  }

  static isAPIKeyAuthButNotInCookie(authScheme: AuthType): boolean {
    return authScheme.type === "apiKey" && authScheme.in !== "cookie";
  }

  static isOAuthWithAuthCodeFlow(authScheme: AuthType): boolean {
    return !!(
      authScheme.type === "oauth2" &&
      authScheme.flows &&
      authScheme.flows.authorizationCode
    );
  }

  static isNotSupportedAuth(authSchemeArray: AuthInfo[][]): boolean {
    if (authSchemeArray.length === 0) {
      return false;
    }

    if (authSchemeArray.length > 0 && authSchemeArray.every((auths) => auths.length > 1)) {
      return true;
    }

    for (const auths of authSchemeArray) {
      if (auths.length === 1) {
        if (
          Utils.isOAuthWithAuthCodeFlow(auths[0].authScheme) ||
          Utils.isBearerTokenAuth(auths[0].authScheme) ||
          Utils.isAPIKeyAuthButNotInCookie(auths[0].authScheme)
        ) {
          return false;
        }
      }
    }

    return true;
  }

  static getAuthArray(
    securities: OpenAPIV3.SecurityRequirementObject[] | undefined,
    spec: OpenAPIV3.Document
  ): AuthInfo[][] {
    const result: AuthInfo[][] = [];
    const securitySchemas = spec.components?.securitySchemes;
    const securitiesArr = securities ?? spec.security;
    if (securitiesArr && securitySchemas) {
      for (let i = 0; i < securitiesArr.length; i++) {
        const security = securitiesArr[i];

        const authArray: AuthInfo[] = [];
        for (const name in security) {
          const auth = securitySchemas[name] as OpenAPIV3.SecuritySchemeObject;
          authArray.push({
            authScheme: auth,
            name: name,
          });
        }

        if (authArray.length > 0) {
          result.push(authArray);
        }
      }
    }

    result.sort((a, b) => a[0].name.localeCompare(b[0].name));

    return result;
  }

  static getAuthMap(spec: OpenAPIV3.Document): OperationAuthInfoMap {
    const authMap: OperationAuthInfoMap = {};

    for (const url in spec.paths) {
      for (const method in spec.paths[url]) {
        const operation = (spec.paths[url] as any)[method] as OpenAPIV3.OperationObject;

        const authArray = Utils.getAuthArray(operation.security, spec);

        if (authArray && authArray.length > 0) {
          const currentAuth = authArray[0][0];
          authMap[operation.operationId!] = currentAuth;
        }
      }
    }

    return authMap;
  }

  static getAuthInfo(spec: OpenAPIV3.Document): AuthInfo | undefined {
    let authInfo: AuthInfo | undefined = undefined;

    for (const url in spec.paths) {
      for (const method in spec.paths[url]) {
        const operation = (spec.paths[url] as any)[method] as OpenAPIV3.OperationObject;

        const authArray = Utils.getAuthArray(operation.security, spec);

        if (authArray && authArray.length > 0) {
          const currentAuth = authArray[0][0];
          if (!authInfo) {
            authInfo = authArray[0][0];
          } else if (authInfo.name !== currentAuth.name) {
            throw new SpecParserError(
              ConstantString.MultipleAuthNotSupported,
              ErrorType.MultipleAuthNotSupported
            );
          }
        }
      }
    }

    return authInfo;
  }

  static updateFirstLetter(str: string): string {
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  static getResponseJson(
    operationObject: OpenAPIV3.OperationObject | undefined,
    allowMultipleMediaType = false
  ): {
    json: OpenAPIV3.MediaTypeObject;
    multipleMediaType: boolean;
  } {
    let json: OpenAPIV3.MediaTypeObject = {};
    let multipleMediaType = false;

    for (const code of ConstantString.ResponseCodeFor20X) {
      const responseObject = operationObject?.responses?.[code] as OpenAPIV3.ResponseObject;
      if (!responseObject) {
        continue;
      }
      multipleMediaType = Utils.containMultipleMediaTypes(responseObject);
      if (!allowMultipleMediaType && multipleMediaType) {
        json = {};
        continue;
      }
      const mediaObj = Utils.getJsonContentType(responseObject);
      if (Object.keys(mediaObj).length > 0) {
        json = mediaObj;
        return { json, multipleMediaType };
      }
    }

    return { json, multipleMediaType };
  }

  static getJsonContentType(
    responseObject: OpenAPIV3.ResponseObject | OpenAPIV3.RequestBodyObject
  ): OpenAPIV3.MediaTypeObject {
    if (responseObject.content) {
      for (const contentType of Object.keys(responseObject.content)) {
        // json media type can also be "application/json; charset=utf-8"
        if (contentType.indexOf("application/json") >= 0) {
          return responseObject.content[contentType];
        }
      }
    }
    return {};
  }

  static convertPathToCamelCase(path: string): string {
    const pathSegments = path.split(/[./{]/);
    const camelCaseSegments = pathSegments.map((segment) => {
      segment = segment.replace(/}/g, "");
      return segment.charAt(0).toUpperCase() + segment.slice(1);
    });
    const camelCasePath = camelCaseSegments.join("");
    return camelCasePath;
  }

  static getUrlProtocol(urlString: string): string | undefined {
    try {
      const url = new URL(urlString);
      return url.protocol;
    } catch (err) {
      return undefined;
    }
  }

  static resolveEnv(str: string): string {
    const placeHolderReg = /\${{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}/g;
    let matches = placeHolderReg.exec(str);
    let newStr = str;
    while (matches != null) {
      const envVar = matches[1];
      const envVal = process.env[envVar];
      if (!envVal) {
        throw new Error(Utils.format(ConstantString.ResolveServerUrlFailed, envVar));
      } else {
        newStr = newStr.replace(matches[0], envVal);
      }
      matches = placeHolderReg.exec(str);
    }
    return newStr;
  }

  static checkServerUrl(servers: OpenAPIV3.ServerObject[], allowHttp = false): ErrorResult[] {
    const errors: ErrorResult[] = [];

    let serverUrl;
    try {
      serverUrl = Utils.resolveEnv(servers[0].url);
    } catch (err) {
      errors.push({
        type: ErrorType.ResolveServerUrlFailed,
        content: (err as Error).message,
        data: servers,
      });
      return errors;
    }

    const protocol = Utils.getUrlProtocol(serverUrl);
    if (!protocol) {
      // Relative server url is not supported
      errors.push({
        type: ErrorType.RelativeServerUrlNotSupported,
        content: ConstantString.RelativeServerUrlNotSupported,
        data: servers,
      });
    } else if (protocol !== "https:" && !(protocol === "http:" && allowHttp)) {
      const protocolString = protocol.slice(0, -1);
      errors.push({
        type: ErrorType.UrlProtocolNotSupported,
        content: Utils.format(ConstantString.UrlProtocolNotSupported, protocol.slice(0, -1)),
        data: protocolString,
      });
    }

    return errors;
  }

  static validateServer(spec: OpenAPIV3.Document, options: ParseOptions): ErrorResult[] {
    const errors: ErrorResult[] = [];

    let hasTopLevelServers = false;
    let hasPathLevelServers = false;
    let hasOperationLevelServers = false;

    const allowHttp = options.projectType === ProjectType.Copilot;

    if (spec.servers && spec.servers.length >= 1) {
      hasTopLevelServers = true;

      // for multiple server, we only use the first url
      const serverErrors = Utils.checkServerUrl(spec.servers, allowHttp);
      errors.push(...serverErrors);
    }

    const paths = spec.paths;
    for (const path in paths) {
      const methods = paths[path];

      if (methods?.servers && methods.servers.length >= 1) {
        hasPathLevelServers = true;
        const serverErrors = Utils.checkServerUrl(methods.servers, allowHttp);

        errors.push(...serverErrors);
      }

      for (const method in methods) {
        const operationObject = (methods as any)[method] as OpenAPIV3.OperationObject;
        if (options.allowMethods?.includes(method) && operationObject) {
          if (operationObject?.servers && operationObject.servers.length >= 1) {
            hasOperationLevelServers = true;
            const serverErrors = Utils.checkServerUrl(operationObject.servers, allowHttp);
            errors.push(...serverErrors);
          }
        }
      }
    }

    if (!hasTopLevelServers && !hasPathLevelServers && !hasOperationLevelServers) {
      errors.push({
        type: ErrorType.NoServerInformation,
        content: ConstantString.NoServerInformation,
      });
    }

    return errors;
  }

  static isWellKnownName(name: string, wellknownNameList: string[]): boolean {
    for (let i = 0; i < wellknownNameList.length; i++) {
      name = name.replace(/_/g, "").replace(/-/g, "");
      if (name.toLowerCase().includes(wellknownNameList[i])) {
        return true;
      }
    }
    return false;
  }

  static generateParametersFromSchema(
    schema: OpenAPIV3.SchemaObject,
    name: string,
    allowMultipleParameters: boolean,
    isRequired = false
  ): [IParameter[], IParameter[]] {
    const requiredParams: IParameter[] = [];
    const optionalParams: IParameter[] = [];

    if (
      schema.type === "string" ||
      schema.type === "integer" ||
      schema.type === "boolean" ||
      schema.type === "number"
    ) {
      const parameter: IParameter = {
        name: name,
        title: Utils.updateFirstLetter(name).slice(0, ConstantString.ParameterTitleMaxLens),
        description: (schema.description ?? "").slice(
          0,
          ConstantString.ParameterDescriptionMaxLens
        ),
      };

      if (allowMultipleParameters) {
        Utils.updateParameterWithInputType(schema, parameter);
      }

      if (isRequired && schema.default === undefined) {
        parameter.isRequired = true;
        requiredParams.push(parameter);
      } else {
        optionalParams.push(parameter);
      }
    } else if (Utils.isObjectSchema(schema)) {
      const { properties } = schema;
      for (const property in properties) {
        let isRequired = false;
        if (schema.required && schema.required?.indexOf(property) >= 0) {
          isRequired = true;
        }
        const [requiredP, optionalP] = Utils.generateParametersFromSchema(
          properties[property] as OpenAPIV3.SchemaObject,
          property,
          allowMultipleParameters,
          isRequired
        );

        requiredParams.push(...requiredP);
        optionalParams.push(...optionalP);
      }
    }

    return [requiredParams, optionalParams];
  }

  static updateParameterWithInputType(schema: OpenAPIV3.SchemaObject, param: IParameter): void {
    if (schema.enum) {
      param.inputType = "choiceset";
      param.choices = [];
      for (let i = 0; i < schema.enum.length; i++) {
        param.choices.push({
          title: schema.enum[i],
          value: schema.enum[i],
        });
      }
    } else if (schema.type === "string") {
      param.inputType = "text";
    } else if (schema.type === "integer" || schema.type === "number") {
      param.inputType = "number";
    } else if (schema.type === "boolean") {
      param.inputType = "toggle";
    }

    if (schema.default) {
      param.value = schema.default;
    }
  }

  static parseApiInfo(
    operationItem: OpenAPIV3.OperationObject,
    options: ParseOptions
  ): IMessagingExtensionCommand {
    const requiredParams: IParameter[] = [];
    const optionalParams: IParameter[] = [];
    const paramObject = operationItem.parameters as OpenAPIV3.ParameterObject[];

    if (paramObject) {
      paramObject.forEach((param: OpenAPIV3.ParameterObject) => {
        const parameter: IParameter = {
          name: param.name,
          title: Utils.updateFirstLetter(param.name).slice(0, ConstantString.ParameterTitleMaxLens),
          description: (param.description ?? "").slice(
            0,
            ConstantString.ParameterDescriptionMaxLens
          ),
        };

        const schema = param.schema as OpenAPIV3.SchemaObject;
        if (options.allowMultipleParameters && schema) {
          Utils.updateParameterWithInputType(schema, parameter);
        }

        if (param.in !== "header" && param.in !== "cookie") {
          if (param.required && schema?.default === undefined) {
            parameter.isRequired = true;
            requiredParams.push(parameter);
          } else {
            optionalParams.push(parameter);
          }
        }
      });
    }

    if (operationItem.requestBody) {
      const requestBody = operationItem.requestBody as OpenAPIV3.RequestBodyObject;
      const requestJson = requestBody.content["application/json"];
      if (Object.keys(requestJson).length !== 0) {
        const schema = requestJson.schema as OpenAPIV3.SchemaObject;
        const [requiredP, optionalP] = Utils.generateParametersFromSchema(
          schema,
          "requestBody",
          !!options.allowMultipleParameters,
          requestBody.required
        );
        requiredParams.push(...requiredP);
        optionalParams.push(...optionalP);
      }
    }

    const operationId = operationItem.operationId!;

    const parameters = [...requiredParams, ...optionalParams];

    const command: IMessagingExtensionCommand = {
      context: ["compose"],
      type: "query",
      title: (operationItem.summary ?? "").slice(0, ConstantString.CommandTitleMaxLens),
      id: operationId,
      parameters: parameters,
      description: (operationItem.description ?? "").slice(
        0,
        ConstantString.CommandDescriptionMaxLens
      ),
    };
    return command;
  }

  static format(str: string, ...args: string[]): string {
    let index = 0;
    return str.replace(/%s/g, () => {
      const arg = args[index++];
      return arg !== undefined ? arg : "";
    });
  }

  static getSafeRegistrationIdEnvName(authName: string): string {
    if (!authName) {
      return "";
    }

    let safeRegistrationIdEnvName = authName.toUpperCase().replace(/[^A-Z0-9_]/g, "_");

    if (!safeRegistrationIdEnvName.match(/^[A-Z]/)) {
      safeRegistrationIdEnvName = "PREFIX_" + safeRegistrationIdEnvName;
    }

    return safeRegistrationIdEnvName;
  }

  static getServerObject(
    spec: OpenAPIV3.Document,
    method: string,
    path: string
  ): OpenAPIV3.ServerObject | undefined {
    const pathObj = spec.paths[path] as any;

    const operationObject = pathObj[method] as OpenAPIV3.OperationObject;

    const rootServer = spec.servers && spec.servers[0];
    const methodServer = spec.paths[path]!.servers && spec.paths[path]!.servers![0];
    const operationServer = operationObject.servers && operationObject.servers[0];

    const serverUrl = operationServer || methodServer || rootServer;

    return serverUrl;
  }

  static getAuthSchemaObject(
    authType: string,
    authParameters: any
  ): OpenAPIV3.SecuritySchemeObject {
    switch (authType) {
      case "oauth":
      case "microsoft-entra":
        return {
          type: "oauth2",
          flows: {
            authorizationCode: {
              authorizationUrl: authParameters.authorizationUrl,
              tokenUrl: authParameters.tokenUrl,
              refreshUrl: authParameters.refreshUrl,
              scopes: authParameters.scopes,
            },
          },
        } as OpenAPIV3.OAuth2SecurityScheme;
      case "api-key":
        return {
          type: "apiKey",
          in: authParameters.in,
          name: authParameters.name,
        } as OpenAPIV3.ApiKeySecurityScheme;
      case "bearer-token":
      default:
        return {
          type: "http",
          scheme: "bearer",
        } as OpenAPIV3.HttpSecurityScheme;
    }
  }
}
