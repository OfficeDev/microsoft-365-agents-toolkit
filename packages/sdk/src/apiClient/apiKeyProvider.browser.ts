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
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "ApiKeyProvider"),
      ErrorCode.RuntimeNotSupported
    );
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
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "ApiKeyProvider"),
        ErrorCode.RuntimeNotSupported
      )
    );
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
