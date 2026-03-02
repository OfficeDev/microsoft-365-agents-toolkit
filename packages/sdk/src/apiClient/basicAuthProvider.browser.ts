// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosRequestConfig } from "axios";
import { AuthProvider } from "./authProvider";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";

/**
 * Provider that handles Basic authentication
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export class BasicAuthProvider implements AuthProvider {
  private userName: string;
  private password: string;

  /**
   *
   * @param { string } userName - Username used in basic auth
   * @param { string } password - Password used in basic auth
   *
   * @throws {@link ErrorCode|InvalidParameter} - when username or password is empty.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  constructor(userName: string, password: string) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "BasicAuthProvider"),
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
   * @throws {@link ErrorCode|AuthorizationInfoAlreadyExists} - when Authorization header or auth property already exists in request configuration.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  public AddAuthenticationInfo(config: AxiosRequestConfig): Promise<AxiosRequestConfig> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "BasicAuthProvider"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}
