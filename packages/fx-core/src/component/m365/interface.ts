// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Not all properties are listed here, only the ones used in the codebase
export interface M365AppDefinition {
  manifestId: string;
  name: string;
  titleId: string;
  version: string;
  scope: string;
  owners: M365AppEntity[];
}

export interface M365AppEntity {
  entityId: string;
  entityType: M365EntityType;
}

export enum M365EntityType {
  User = "User",
  Group = "Group",
}
