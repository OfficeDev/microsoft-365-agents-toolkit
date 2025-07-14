// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken, GetTokenOptions, TokenCredential } from "@azure/identity";
import { UserInfo } from "../models/userinfo";
import { OnBehalfOfCredentialAuthConfig } from "../models/configuration";
import { formatString } from "../util/utils";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";

/**
 * Represent on-behalf-of flow to get user identity, and it is designed to be used in Azure Function or Bot scenarios.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @remarks
 * Can only be used in server side.
 */
export class OnBehalfOfUserCredential implements TokenCredential {
  /**
   * Constructor of OnBehalfOfUserCredential
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @remarks
   * Can Only works in in server side.
   */
  constructor(ssoToken: string, config: OnBehalfOfCredentialAuthConfig) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "OnBehalfOfUserCredential"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Get access token from credential.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @remarks
   * Can only be used in server side.
   */
  getToken(scopes: string | string[], options?: GetTokenOptions): Promise<AccessToken | null> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "OnBehalfOfUserCredential"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Get basic user info from SSO token.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @remarks
   * Can only be used in server side.
   */
  public getUserInfo(): Promise<UserInfo> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "OnBehalfOfUserCredential"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}
