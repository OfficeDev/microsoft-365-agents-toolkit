// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface ResourcePermission {
  name: string;
  type: string;
  resourceId: string | undefined;
  roles: string[] | undefined;
}

export enum CollaborationState {
  OK = "OK",
  NotProvisioned = "NotProvisioned",
  M365TenantNotMatch = "M365TenantNotMatch",
  EmptyM365Tenant = "EmptyM365Tenant",
  m365AccountNotSignedIn = "M365AccountNotSignedIn",
}

export interface CollaborationStateResult {
  state: CollaborationState;
  message?: string;
}

export interface ListCollaboratorResult {
  state: CollaborationState;
  message?: string;
  collaborators?: Collaborator[];
  error?: any;
}

export interface PermissionsResult {
  state: CollaborationState;
  message?: string;
  userInfo?: Record<string, any>;
  permissions?: ResourcePermission[];
}

export interface Collaborator {
  userPrincipalName: string;
  userObjectId: string;
  isAadOwner: boolean;
  teamsAppResourceId: string;
  aadResourceId?: string;
}

export interface AadOwner {
  userObjectId: string;
  resourceId: string;
  displayName: string;
  userPrincipalName: string;
}

export interface TeamsAppAdmin {
  userObjectId: string;
  resourceId: string;
  displayName: string;
  userPrincipalName: string;
}

export interface AgentOwner {
  userObjectId: string;
  resourceId: string;
  displayName: string;
  userPrincipalName: string;
}

export interface AppIds {
  teamsAppId?: string;
  aadObjectId?: string;
}
