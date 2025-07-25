// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface AppUser {
  tenantId: string;
  aadId: string;
  displayName: string;
  userPrincipalName: string;
  isAdministrator: boolean;
}
