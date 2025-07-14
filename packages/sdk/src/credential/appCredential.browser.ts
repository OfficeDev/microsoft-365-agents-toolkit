// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken, TokenCredential, GetTokenOptions } from "@azure/identity";
import { AppCredentialAuthConfig } from "../models/configuration";
import { formatString } from "../util/utils";
import { ErrorCode, ErrorMessage, ErrorWithCode } from "../core/errors";

/**
 * Represent Microsoft 365 tenant identity, and it is usually used when user is not involved.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @remarks
 * Only works in in server side.
 */
export class AppCredential implements TokenCredential {
  /**
   * Constructor of AppCredential.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @remarks
   * Only works in in server side.
   */
  constructor(authConfig: AppCredentialAuthConfig) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "AppCredential"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Get access token for credential.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @remarks
   * Only works in in server side.
   */
  getToken(scopes: string | string[], options?: GetTokenOptions): Promise<AccessToken | null> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "AppCredential"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}
