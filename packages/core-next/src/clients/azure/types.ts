// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Azure Resource Manager constants and types.
 */

/** Azure management endpoint for public cloud */
export const ARM_BASE_URL = "https://management.azure.com";

/** ARM API version for resource deployments */
export const ARM_DEPLOYMENT_API_VERSION = "2021-04-01";

/** ARM API version for web site management */
export const ARM_WEBAPPS_API_VERSION = "2024-04-01";

/** Default scope for Azure management operations */
export function azureManagementScopes(): string[] {
  return [`${ARM_BASE_URL}/.default`];
}

/** Deployment mode for ARM templates */
export type DeploymentMode = "Incremental" | "Complete";

/** ARM deployment provisioning states */
export type ProvisioningState =
  | "Accepted"
  | "Running"
  | "Ready"
  | "Creating"
  | "Created"
  | "Deleting"
  | "Deleted"
  | "Canceled"
  | "Failed"
  | "Succeeded"
  | "Updating";

/** Terminal provisioning states that end polling */
const TERMINAL_STATES: ReadonlySet<string> = new Set(["Succeeded", "Failed", "Canceled"]);

export function isTerminalState(state: string): boolean {
  return TERMINAL_STATES.has(state);
}

/** ARM template deployment request body */
export interface ArmDeploymentRequest {
  properties: {
    template: Record<string, unknown>;
    parameters?: Record<string, unknown> | null;
    mode: DeploymentMode;
  };
}

/** ARM deployment response (subset) */
export interface ArmDeploymentResponse {
  id?: string;
  name?: string;
  properties?: {
    provisioningState?: string;
    outputs?: Record<string, { type: string; value: unknown }>;
    error?: ArmDeploymentError;
    timestamp?: string;
  };
}

/** ARM deployment error shape */
export interface ArmDeploymentError {
  code?: string;
  message?: string;
  details?: ArmDeploymentError[];
}

/** ARM template configuration (one template to deploy) */
export interface ArmTemplateConfig {
  /** Path to .bicep or .json template file */
  path: string;
  /** Path to .json parameters file (optional) */
  parameters?: string;
  /** Azure deployment name */
  deploymentName: string;
}

/** Result of parsing an Azure resource ID */
export interface AzureResourceId {
  subscriptionId: string;
  resourceGroupName: string;
  instanceId: string;
}

/** Zip deployment status values */
export enum DeployStatus {
  Building = 0,
  Deploying = 1,
  Pending = 2,
  Failed = 3,
  Success = 4,
}
