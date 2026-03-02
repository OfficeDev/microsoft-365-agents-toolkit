// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * UserInfo with user displayName, objectId and preferredUserName.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export interface UserInfo {
  /**
   * User Display Name.
   *
   * @readonly
   */
  displayName: string;

  /**
   * User unique reference within the Azure Active Directory domain.
   *
   * @readonly
   */
  objectId: string;

  /**
   * User tenant ID.
   *
   * @readonly
   */
  tenantId: string;

  /**
   * Usually be the email address.
   *
   * @readonly
   */
  preferredUserName: string;
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export interface UserTenantIdAndLoginHint {
  /**
   * User Login Hint.
   *
   * @readonly
   */
  loginHint: string;

  /**
   * User Tenant Id.
   *
   * @readonly
   */
  tid: string;
}
