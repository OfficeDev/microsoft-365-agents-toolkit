// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export interface AuthCodeResult {
  code: string;
  codeVerifier: string;
  redirectUri: string;
}
