// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosRequestConfig } from "axios";
import { AuthProvider } from "./authProvider";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";

/**
 * Provider that handles API Key authentication
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export class ApiKeyProvider implements AuthProvider {
  private keyName: string;
  private keyValue: string;
  private keyLocation: ApiKeyLocation;

  /**
   *
   * @param { string } keyName - The name of request header or query parameter that specifies API Key
   * @param { string } keyValue - The value of API Key
   * @param { ApiKeyLocation } keyLocation - The location of API Key: request header or query parameter.
   *
   * @throws {@link ErrorCode|InvalidParameter} - when key name or key value is empty.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  constructor(keyName: string, keyValue: string, keyLocation: ApiKeyLocation) {
    if (!keyName) {
      throw new ErrorWithCode(
        formatString(ErrorMessage.EmptyParameter, "keyName"),
        ErrorCode.InvalidParameter
      );
    }
    if (!keyValue) {
      throw new ErrorWithCode(
        formatString(ErrorMessage.EmptyParameter, "keyVaule"),
        ErrorCode.InvalidParameter
      );
    }
    this.keyName = keyName;
    this.keyValue = keyValue;
    this.keyLocation = keyLocation;
  }

  /**
   * Adds authentication info to http requests
   *
   * @param { AxiosRequestConfig } config - Contains all the request information and can be updated to include extra authentication info.
   * Refer https://axios-http.com/docs/req_config for detailed document.
   *
   * @returns Updated axios request config.
   *
   * @throws {@link ErrorCode|AuthorizationInfoAlreadyExists} - when API key already exists in request header or url query parameter.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  public AddAuthenticationInfo(config: AxiosRequestConfig): Promise<AxiosRequestConfig> {
    switch (this.keyLocation) {
      case ApiKeyLocation.Header:
        if (!config.headers) {
          config.headers = {};
        }
        if (config.headers[this.keyName]) {
          return Promise.reject(
            new ErrorWithCode(
              formatString(ErrorMessage.DuplicateApiKeyInHeader, this.keyName),
              ErrorCode.AuthorizationInfoAlreadyExists
            )
          );
        }
        config.headers[this.keyName] = this.keyValue;
        break;
      case ApiKeyLocation.QueryParams:
        if (!config.params) {
          config.params = {};
        }
        let urlHasDefinedApiKey = false;
        if (config.url) {
          const url = new URL(config.url, config.baseURL);
          urlHasDefinedApiKey = url.searchParams.has(this.keyName);
        }
        if (config.params[this.keyName] || urlHasDefinedApiKey) {
          return Promise.reject(
            new ErrorWithCode(
              formatString(ErrorMessage.DuplicateApiKeyInQueryParam, this.keyName),
              ErrorCode.AuthorizationInfoAlreadyExists
            )
          );
        }
        config.params[this.keyName] = this.keyValue;
        break;
    }

    return Promise.resolve(config);
  }
}

/**
 * Define available location for API Key location
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export enum ApiKeyLocation {
  /**
   * The API Key is placed in request header
   */
  Header,
  /**
   * The API Key is placed in query parameter
   */
  QueryParams,
}
