// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Not all properties are listed here, only the ones used in the codebase
export interface M365AppDefinition {
  manifestId: string;
  name: string;
  titleId: string;
  version: string;
  scope: string;
  owners: M365AppOwners[];
}

export interface M365AppOwners {
  entityId: string;
  entityType: M365OwnerType;
}

export enum M365OwnerType {
  User = "user",
}
