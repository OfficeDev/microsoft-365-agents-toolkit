// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosRequestConfig } from "axios";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { AuthProvider } from "./authProvider";

/**
 * Provider that handles Bearer Token authentication
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export class BearerTokenAuthProvider implements AuthProvider {
  private getToken: () => Promise<string>;

  /**
   * @param { () => Promise<string> } getToken - Function that returns the content of bearer token used in http request
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  constructor(getToken: () => Promise<string>) {
    this.getToken = getToken;
  }

  /**
   * Adds authentication info to http requests
   *
   * @param { AxiosRequestConfig } config - Contains all the request information and can be updated to include extra authentication info.
   * Refer https://axios-http.com/docs/req_config for detailed document.
   *
   * @returns Updated axios request config.
   *
   * @throws {@link ErrorCode|AuthorizationInfoAlreadyExists} - when Authorization header already exists in request configuration.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  public async AddAuthenticationInfo(config: AxiosRequestConfig): Promise<AxiosRequestConfig> {
    const token = await this.getToken();
    if (!config.headers) {
      config.headers = {};
    }
    if (config.headers["Authorization"]) {
      throw new ErrorWithCode(
        ErrorMessage.AuthorizationHeaderAlreadyExists,
        ErrorCode.AuthorizationInfoAlreadyExists
      );
    }

    config.headers["Authorization"] = `Bearer ${token}`;
    return config;
  }
}
