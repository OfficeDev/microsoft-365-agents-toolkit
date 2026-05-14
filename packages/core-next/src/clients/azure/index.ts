// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export { AzureArmClient } from "./client";
export {
  ARM_BASE_URL,
  ARM_DEPLOYMENT_API_VERSION,
  ARM_WEBAPPS_API_VERSION,
  azureManagementScopes,
  isTerminalState,
  DeployStatus,
} from "./types";
export type {
  ArmDeploymentRequest,
  ArmDeploymentResponse,
  ArmDeploymentError,
  ArmTemplateConfig,
  AzureResourceId,
  DeploymentMode,
  ProvisioningState,
} from "./types";
