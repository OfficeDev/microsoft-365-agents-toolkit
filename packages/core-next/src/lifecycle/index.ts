// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type {
  DriverStep,
  LifecycleName,
  RawProjectModel,
  ProjectModel,
  StepResult,
  LifecycleResult,
  LifecycleProgress,
  LifecycleAnalysis,
  M365TenantInfo,
  AzureAccountInfo,
  SubscriptionInfo,
  ResourceGroupInfo,
  PostAction,
  LifecycleOperationResult,
} from "./types";
export { LIFECYCLE_NAMES } from "./types";

export { parseProjectYaml } from "./parser";
export { resolveConfig } from "./resolver";
export { executeLifecycle } from "./executor";
export { createProgressAdapter, silentProgress } from "./progress";
export { analyzeSteps } from "./analyze";
export {
  ensureM365Auth,
  ensureAzureAuth,
  ensureSubscription,
  ensureResourceGroup,
  ensureResourceSuffix,
  confirmProvision,
  confirmDeploy,
} from "./prerequisites";
export { provisionOp, deployOp, publishOp } from "./operations";
