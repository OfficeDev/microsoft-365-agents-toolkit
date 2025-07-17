// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { SSOTokenInfoBase, SSOTokenV1Info, SSOTokenV2Info } from "../models/ssoTokenInfo";
import { UserInfo, UserTenantIdAndLoginHint } from "../models/userinfo";
import { jwtDecode } from "jwt-decode";
import { internalLogger } from "./logger";
import { AccessToken } from "@azure/identity";
import { AuthenticationResult } from "@azure/msal-browser";

/**
 * Parse jwt token payload
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @param token
 *
 * @returns Payload object
 *
 * @internal
 */
export function parseJwt(token: string): SSOTokenInfoBase {
  try {
    const tokenObj: SSOTokenInfoBase = jwtDecode(token);
    if (!tokenObj || !tokenObj.exp) {
      throw new ErrorWithCode(
        "Decoded token is null or exp claim does not exists.",
        ErrorCode.InternalError
      );
    }

    return tokenObj;
  } catch (err: any) {
    const errorMsg = "Parse jwt token failed in node env with error: " + (err.message as string);
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.InternalError);
  }
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function getUserInfoFromSsoToken(ssoToken: string): UserInfo {
  if (!ssoToken) {
    const errorMsg = "SSO token is undefined.";
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.InvalidParameter);
  }
  const tokenObject = parseJwt(ssoToken) as SSOTokenV1Info | SSOTokenV2Info;

  const userInfo: UserInfo = {
    displayName: tokenObject.name,
    objectId: tokenObject.oid,
    tenantId: tokenObject.tid,
    preferredUserName: "",
  };

  if (tokenObject.ver === "2.0") {
    userInfo.preferredUserName = (tokenObject as SSOTokenV2Info).preferred_username;
  } else if (tokenObject.ver === "1.0") {
    userInfo.preferredUserName = (tokenObject as SSOTokenV1Info).upn;
  }
  return userInfo;
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function getTenantIdAndLoginHintFromSsoToken(ssoToken: string): UserTenantIdAndLoginHint {
  if (!ssoToken) {
    const errorMsg = "SSO token is undefined.";
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.InvalidParameter);
  }
  const tokenObject = parseJwt(ssoToken) as SSOTokenV1Info | SSOTokenV2Info;

  const userInfo: UserTenantIdAndLoginHint = {
    tid: tokenObject.tid,
    loginHint:
      tokenObject.ver === "2.0"
        ? (tokenObject as SSOTokenV2Info).preferred_username
        : (tokenObject as SSOTokenV1Info).upn,
  };

  return userInfo;
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function parseAccessTokenFromAuthCodeTokenResponse(
  tokenResponse: string | AuthenticationResult
): AccessToken {
  try {
    const tokenResponseObject =
      typeof tokenResponse == "string"
        ? (JSON.parse(tokenResponse) as AuthenticationResult)
        : tokenResponse;
    if (!tokenResponseObject || !tokenResponseObject.accessToken) {
      const errorMsg = "Get empty access token from Auth Code token response.";

      internalLogger.error(errorMsg);
      throw new Error(errorMsg);
    }

    const token = tokenResponseObject.accessToken;
    const tokenObject = parseJwt(token);

    if (tokenObject.ver !== "1.0" && tokenObject.ver !== "2.0") {
      const errorMsg = "SSO token is not valid with an unknown version: " + tokenObject.ver;
      internalLogger.error(errorMsg);
      throw new Error(errorMsg);
    }

    const accessToken: AccessToken = {
      token: token,
      expiresOnTimestamp: tokenObject.exp * 1000,
    };
    return accessToken;
  } catch (error: any) {
    const errorMsg =
      "Parse access token failed from Auth Code token response in node env with error: " +
      (error.message as string);
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.InternalError);
  }
}

/**
 * Format string template with replacements
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * ```typescript
 * const template = "{0} and {1} are fruit. {0} is my favorite one."
 * const formattedStr = formatString(template, "apple", "pear"); // formattedStr: "apple and pear are fruit. apple is my favorite one."
 * ```
 *
 * @param str string template
 * @param replacements replacement string array
 * @returns Formatted string
 *
 * @internal
 */
export function formatString(str: string, ...replacements: string[]): string {
  const args = replacements;
  return str.replace(/{(\d+)}/g, function (match, number) {
    return typeof args[number] != "undefined" ? args[number] : match;
  });
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function validateScopesType(value: any): void {
  // string
  if (typeof value === "string" || value instanceof String) {
    return;
  }

  // empty array
  if (Array.isArray(value) && value.length === 0) {
    return;
  }

  // string array
  if (Array.isArray(value) && value.length > 0 && value.every((item) => typeof item === "string")) {
    return;
  }

  const errorMsg = "The type of scopes is not valid, it must be string or string array";
  internalLogger.error(errorMsg);
  throw new ErrorWithCode(errorMsg, ErrorCode.InvalidParameter);
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function getScopesArray(scopes: string | string[]): string[] {
  const scopesArray: string[] = typeof scopes === "string" ? scopes.split(" ") : scopes;
  return scopesArray.filter((x) => x !== null && x !== "");
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function getAuthority(authorityHost: string, tenantId: string): string {
  const normalizedAuthorityHost = authorityHost.replace(/\/+$/g, "");
  return normalizedAuthorityHost + "/" + tenantId;
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function validateConfig(config: any): void {
  if (
    config.clientId &&
    (config.clientSecret || config.certificateContent) &&
    config.tenantId &&
    config.authorityHost
  ) {
    return;
  }

  const missingValues = [];

  if (!config.clientId) {
    missingValues.push("clientId");
  }

  if (!config.clientSecret && !config.certificateContent) {
    missingValues.push("clientSecret or certificateContent");
  }

  if (!config.tenantId) {
    missingValues.push("tenantId");
  }

  if (!config.authorityHost) {
    missingValues.push("authorityHost");
  }

  const errorMsg = formatString(
    ErrorMessage.InvalidConfiguration,
    missingValues.join(", "),
    "undefined"
  );
  internalLogger.error(errorMsg);
  throw new ErrorWithCode(errorMsg, ErrorCode.InvalidConfiguration);
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export interface ClientCertificate {
  thumbprintSha256: string;
  privateKey: string;
}
