// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface AppUser {
  tenantId: string;
  aadId: string;
  displayName: string;
  userPrincipalName: string;
  isAdministrator: boolean;
}

export interface AppGroup {
  id: string;
  displayName: string;
  email: string;
}
