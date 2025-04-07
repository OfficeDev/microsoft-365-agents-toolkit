// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import axios, {
  AxiosInstance,
  AxiosError,
  CreateAxiosDefaults,
  AxiosResponse,
  InternalAxiosRequestConfig,
} from "axios";
import { TOOLS } from "./globalVars";
import { APP_STUDIO_API_NAMES, Constants } from "../component/driver/teamsApp/constants";
import {
  TelemetryPropertyKey,
  TelemetryPropertyValue,
} from "../component/driver/teamsApp/utils/telemetry";
import { TelemetryEvent, TelemetryProperty, TelemetrySuccess } from "./telemetry";
import { DeveloperPortalAPIFailedSystemError } from "../error/teamsApp";
import { HttpMethod } from "../component/constant/commonConstant";
import { getDefaultString } from "./localizeUtils";
import { MOS3Api, MOS3ApiDefinitions } from "../component/m365/serviceConstant";

/**
 * This client will send telemetries to record API request trace
 */
export class WrappedAxiosClient {
  public static create(config?: CreateAxiosDefaults): AxiosInstance {
    const instance = axios.create(config);

    instance.interceptors.request.use((request) => this.onRequest(request));

    // eslint-disable-next-line prettier/prettier
    instance.interceptors.response.use((response) => this.onResponse(response), (error) => this.onRejected(error));

    return instance;
  }

  /**
   * Send API start telemetry
   * @param request
   */
  public static onRequest(request: InternalAxiosRequestConfig) {
    const method = request.method!;
    const fullPath = `${request.baseURL ?? ""}${request.url ?? ""}`;
    const apiName = this.convertUrlToApiName(fullPath, method);

    const properties: { [key: string]: string } = {
      url: `<${apiName}-url>`,
      method: method,
      params: this.generateParameters(request.params),
      ...this.generateExtraProperties(fullPath, request.data),
    };
    const eventName = this.getEventName(fullPath);
    TOOLS?.telemetryReporter?.sendTelemetryEvent(`${eventName}-start`, properties);
    return request;
  }

  /**
   * Send API success telemetry
   * @param response
   * @returns
   */
  public static onResponse(response: AxiosResponse) {
    const method = response.request.method;
    const fullPath = `${(response.request.host as string) ?? ""}${
      (response.request.path as string) ?? ""
    }`;
    const apiName = this.convertUrlToApiName(fullPath, method);

    const properties: { [key: string]: string } = {
      url: `<${apiName}-url>`,
      method: method,
      params: this.generateParameters(response.config.params),
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      "status-code": response.status.toString(),
      ...this.generateExtraProperties(fullPath, response.data),
    };

    const eventName = this.getEventName(fullPath);
    TOOLS?.telemetryReporter?.sendTelemetryEvent(eventName, properties);
    return response;
  }

  /**
   * Send API failure telemetry
   * @param error
   * @returns
   */
  public static onRejected(error: AxiosError) {
    const method = error.request.method as string;
    const fullPath = `${(error.request.host as string) ?? ""}${
      (error.request.path as string) ?? ""
    }`;
    const apiName = this.convertUrlToApiName(fullPath, method);

    let requestData: any;
    if (error.config?.data && typeof error.config.data === "string") {
      try {
        requestData = JSON.parse(error.config.data);
      } catch (error) {
        requestData = undefined;
      }
    }
    const properties: { [key: string]: string } = {
      url: `<${apiName}-url>`,
      method: method,
      params: this.generateParameters(error.config!.params),
      [TelemetryProperty.Success]: TelemetrySuccess.No,
      [TelemetryProperty.ErrorMessage]: error.response
        ? JSON.stringify(error.response.data)
        : error.message ?? "undefined",
      "status-code": error.response?.status.toString() ?? "undefined",
      ...this.generateExtraProperties(fullPath, requestData),
    };

    const eventName = this.getEventName(fullPath);
    if (eventName === TelemetryEvent.AppStudioApi) {
      const correlationId = error.response?.headers[Constants.CORRELATION_ID] ?? "undefined";
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      const extraData = getDefaultString(
        "error.appstudio.apiFailed.reason.common",
        error.response?.data ? `data: ${JSON.stringify(error.response.data)}` : ""
      );
      const TDPApiFailedError = new DeveloperPortalAPIFailedSystemError(
        error,
        correlationId,
        apiName,
        extraData
      );
      properties[
        TelemetryProperty.ErrorCode
      ] = `${TDPApiFailedError.source}.${TDPApiFailedError.name}`;
      properties[TelemetryProperty.ErrorMessage] = TDPApiFailedError.message;
      properties[TelemetryProperty.TDPTraceId] = correlationId;
    } else if (eventName === TelemetryEvent.MOSApi) {
      const tracingId = (error.response?.headers?.traceresponse ?? "undefined") as string;
      const originalMessage = error.message;
      const innerError = (error.response?.data as any).error || { code: "", message: "" };
      const finalMessage = `${originalMessage} (tracingId: ${tracingId}) ${
        innerError.code as string
      }: ${innerError.message as string} `;
      properties[TelemetryProperty.ErrorMessage] = finalMessage;
      properties[TelemetryProperty.MOSTraceId] = tracingId;
      const relativePath = (error.request.path || "") as string;
      const mosApiDef = this.convertMethodUrlToApiDefForMOS(method, relativePath);
      if (mosApiDef) {
        properties.url = `mos_${mosApiDef.key}`;
      }
    }

    TOOLS?.telemetryReporter?.sendTelemetryErrorEvent(eventName, properties);
    return Promise.reject(error);
  }

  static convertMethodUrlToApiDefForMOS(method: string, url: string): MOS3Api | undefined {
    for (const key of Object.keys(MOS3ApiDefinitions)) {
      const api = MOS3ApiDefinitions[key];
      if (api.method === method && url.match(api.path)) {
        return api;
      }
    }
    return undefined;
  }

  /**
   * Convert request URL to API name, otherwise it will be redacted in telemetry
   * This function should be extended when new API is added
   * @param baseUrl
   * @param path
   * @param method
   * @returns
   */
  public static convertUrlToApiName(fullPath: string, method: string): string {
    if (this.isTDPApi(fullPath)) {
      if (fullPath.match(new RegExp("/api/appdefinitions/partnerCenterAppPackageValidation"))) {
        return APP_STUDIO_API_NAMES.VALIDATE_APP_PACKAGE;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/v2/import"))) {
        return APP_STUDIO_API_NAMES.CREATE_APP;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/manifest"))) {
        return APP_STUDIO_API_NAMES.EXISTS_IN_TENANTS;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/.*/manifest"))) {
        return APP_STUDIO_API_NAMES.GET_APP_PACKAGE;
      }
      if (fullPath.match(new RegExp("/api/appdefinitions/.*/owner"))) {
        return APP_STUDIO_API_NAMES.UPDATE_OWNER;
      }
      if (
        fullPath.match(
          new RegExp(
            /\/api\/appdefinitions\/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/
          )
        )
      ) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_APP;
        }
        if (method.toUpperCase() === HttpMethod.DELETE) {
          return APP_STUDIO_API_NAMES.DELETE_APP;
        }
      }
      if (fullPath.match(new RegExp("/api/appdefinitions"))) {
        return APP_STUDIO_API_NAMES.LIST_APPS;
      }
      if (fullPath.match(new RegExp("/api/publishing/.*/appdefinitions"))) {
        return APP_STUDIO_API_NAMES.UPDATE_PUBLISHED_APP;
      }
      if (fullPath.match(new RegExp("/api/publishing/.*"))) {
        return APP_STUDIO_API_NAMES.GET_PUBLISHED_APP;
      }
      if (fullPath.match(new RegExp("/api/publishing"))) {
        return APP_STUDIO_API_NAMES.PUBLISH_APP;
      }
      if (fullPath.match(new RegExp("/api/usersettings/mtUserAppPolicy"))) {
        return APP_STUDIO_API_NAMES.CHECK_SIDELOADING_STATUS;
      }
      if (fullPath.match(new RegExp("/api/v1.0/apiSecretRegistrations/.*"))) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_API_KEY;
        }
        if (method.toUpperCase() === HttpMethod.PATCH) {
          return APP_STUDIO_API_NAMES.UPDATE_API_KEY;
        }
      }
      if (fullPath.match(new RegExp("/api/v1.0/apiSecretRegistrations"))) {
        return APP_STUDIO_API_NAMES.CREATE_API_KEY;
      }
      if (
        fullPath.match(
          new RegExp(
            /\/api\/botframework\/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/
          )
        )
      ) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_BOT;
        }
        if (method.toUpperCase() === HttpMethod.POST) {
          return APP_STUDIO_API_NAMES.UPDATE_BOT;
        }
        if (method.toUpperCase() === HttpMethod.DELETE) {
          return APP_STUDIO_API_NAMES.DELETE_BOT;
        }
      }
      if (fullPath.match(new RegExp("/api/botframework"))) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.LIST_BOT;
        }
        if (method.toUpperCase() === HttpMethod.POST) {
          return APP_STUDIO_API_NAMES.CREATE_BOT;
        }
      }
      if (fullPath.match(new RegExp("/api/v1.0/appvalidations/appdefinition/validate"))) {
        return APP_STUDIO_API_NAMES.SUBMIT_APP_VALIDATION;
      }
      if (
        fullPath.match(
          new RegExp(
            "/api/v1.0/appvalidations/appdefinitions/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
          )
        )
      ) {
        return APP_STUDIO_API_NAMES.GET_APP_VALIDATION_REQUESTS;
      }
      if (
        fullPath.match(
          new RegExp(
            "/api/v1.0/appvalidations/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
          )
        )
      ) {
        return APP_STUDIO_API_NAMES.GET_APP_VALIDATION_RESULT;
      }
      if (fullPath.match(new RegExp("/api/v1.0/oAuthConfigurations/.*"))) {
        if (method.toUpperCase() === HttpMethod.GET) {
          return APP_STUDIO_API_NAMES.GET_OAUTH;
        }
        if (method.toUpperCase() === HttpMethod.PATCH) {
          return APP_STUDIO_API_NAMES.UPDATE_OAUTH;
        }
      }
      if (fullPath.match(new RegExp("/api/v1.0/oAuthConfigurations"))) {
        return APP_STUDIO_API_NAMES.CREATE_OAUTH;
      }
    }
    if (
      fullPath.match(
        new RegExp(/(^https:\/\/)?authsvc\.teams\.microsoft\.com\/v1\.0\/users\/region/)
      )
    ) {
      return "get-region";
    }
    return fullPath.replace(/\//g, `-`);
  }

  /**
   * Generate extra properties for specific requirements
   * @param baseUrl
   * @param path
   * @param method
   */
  private static generateExtraProperties(fullPath: string, data?: any): { [key: string]: string } {
    const properties: { [key: string]: string } = {};
    if (this.isTDPApi(fullPath)) {
      // Add region property
      properties[TelemetryPropertyKey.region] = String(this.extractRegion(fullPath));

      // Add bot id property
      if (fullPath.match(new RegExp("/api/botframework"))) {
        const regex = new RegExp(/\/api\/botframework\/([0-9a-fA-F-]+)/);
        const matches = regex.exec(fullPath);
        if (matches != null && matches.length > 1) {
          properties[TelemetryProperty.BotId] = matches[1];
        } else if (data?.botId) {
          properties[TelemetryProperty.BotId] = data.botId;
        }
      }
    }
    return properties;
  }

  /**
   * Extract region from baseUrl, E.g. https://dev.teams.microsoft.com/amer => amer
   * @param url
   * @returns
   */
  private static extractRegion(fullPath: string): string | undefined {
    const regex = /dev(-int)?\.teams\.microsoft\.com\/([a-zA-Z-_]+)\/api/;
    const matches = regex.exec(fullPath);
    if (matches != null && matches.length > 1) {
      return matches[2];
    }

    return TelemetryPropertyValue.Global;
  }

  /**
   * Check if it's TDP Api
   * @param baseUrl
   * @returns
   */
  private static isTDPApi(baseUrl: string): boolean {
    const regex = /(^https:\/\/)?dev(-int)?\.teams\.microsoft\.com/;
    const matches = regex.exec(baseUrl);
    return matches != null && matches.length > 0;
  }

  private static getEventName(
    baseUrl: string
  ): TelemetryEvent.MOSApi | TelemetryEvent.AppStudioApi | TelemetryEvent.DependencyApi {
    if (this.isTDPApi(baseUrl)) {
      return TelemetryEvent.AppStudioApi;
    } else if (baseUrl.includes("titles.prod.mos.microsoft.com")) {
      return TelemetryEvent.MOSApi;
    } else {
      return TelemetryEvent.DependencyApi;
    }
  }

  /**
   * Flattern query parameters to string, e.g. {a: 1, b: 2} => a:1;b:2
   * @param params
   * @returns
   */
  private static generateParameters(params?: Record<string, unknown>): string {
    if (!params) {
      return "";
    }
    const parameters: string[] = [];
    for (const [key, value] of Object.entries(params)) {
      parameters.push(`${key}:${value as string}`);
    }
    return parameters.join(";");
  }
}
