// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Base class for SSO token payload
 * Payload claims can be found here: https://docs.microsoft.com/en-us/azure/active-directory/develop/id-tokens
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export interface SSOTokenInfoBase {
  aud: string;
  iss: string;
  iat: number;
  nbf: number;
  exp: number;
  aio: string;
  name: string;
  oid: string;
  rh: string;
  scp: string;
  sub: string;
  tid: string;
  uti: string;
  ver: string;
}

/**
 * SSO token v1
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export interface SSOTokenV2Info extends SSOTokenInfoBase {
  azp: string;
  azpacr: string;
  preferred_username: string;
}

/**
 * SSO token v2
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export interface SSOTokenV1Info extends SSOTokenInfoBase {
  acr: string;
  amr: string[];
  appid: string;
  appidacr: string;
  family_name: string;
  given_name: string;
  ipaddr: string;
  onprem_sid: string;
  unique_name: string;
  upn: string;
}
