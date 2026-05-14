// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Lifecycle engine types.
 *
 * Preserves the same m365agents.yml schema as fx-core v3 while using
 * the v4 DriverRegistry for driver lookup.
 */

/**
 * A single driver step defined in the YAML lifecycle.
 */
export interface DriverStep {
  /** Optional human-readable step name. */
  name?: string;
  /** Driver ID (e.g. "teamsApp/create", "arm/deploy"). */
  uses: string;
  /** Driver-specific configuration. */
  with: Record<string, unknown>;
  /** Optional environment variables injected before execution. */
  env?: Record<string, string>;
  /** Optional mapping of driver outputs to environment file entries. */
  writeToEnvironmentFile?: Record<string, string>;
}

/**
 * Lifecycle names recognized in the YAML file.
 */
export const LIFECYCLE_NAMES = [
  "registerApp",
  "configureApp",
  "provision",
  "deploy",
  "publish",
  "share",
] as const;

export type LifecycleName = (typeof LIFECYCLE_NAMES)[number];

/**
 * Raw project model parsed from YAML (before driver resolution).
 */
export interface RawProjectModel {
  version: string;
  environmentFolderPath?: string;
  additionalMetadata?: Record<string, unknown>;
  registerApp?: DriverStep[];
  configureApp?: DriverStep[];
  provision?: DriverStep[];
  deploy?: DriverStep[];
  publish?: DriverStep[];
  share?: DriverStep[];
}

/**
 * Resolved project model ready for execution.
 * Each lifecycle section is an ordered list of resolved steps.
 */
export interface ProjectModel {
  version: string;
  environmentFolderPath?: string;
  additionalMetadata?: Record<string, unknown>;
  registerApp?: DriverStep[];
  configureApp?: DriverStep[];
  provision?: DriverStep[];
  deploy?: DriverStep[];
  publish?: DriverStep[];
  share?: DriverStep[];
}

/**
 * Result of executing a single driver step.
 */
export interface StepResult {
  /** Driver ID that was executed. */
  driver: string;
  /** Key-value outputs produced by the driver. */
  outputs: Record<string, string>;
  /** Time taken in milliseconds. */
  durationMs: number;
}

/**
 * Result of executing an entire lifecycle.
 */
export interface LifecycleResult {
  /** Lifecycle name (e.g. "provision"). */
  lifecycle: LifecycleName;
  /** Results for each step in order. */
  steps: StepResult[];
  /** Total time in milliseconds. */
  totalDurationMs: number;
}

// ---------------------------------------------------------------------------
// Progress reporting
// ---------------------------------------------------------------------------

/**
 * Callback interface for lifecycle progress reporting.
 *
 * Operations and the executor call these hooks at each stage.
 * The consumer (VS Code, CLI) provides an implementation that maps to
 * their native progress UI. A default adapter is available in progress.ts.
 */
export interface LifecycleProgress {
  /** Called when the lifecycle begins. */
  onStart(lifecycle: LifecycleName, totalSteps: number): void | Promise<void>;
  /** Called before each driver step executes. */
  onStepStart(stepIndex: number, stepLabel: string): void | Promise<void>;
  /** Called after each driver step completes. */
  onStepComplete(stepIndex: number, stepLabel: string, durationMs: number): void | Promise<void>;
  /** Called when the lifecycle finishes (success or failure). */
  onEnd(success: boolean): void | Promise<void>;
}

// ---------------------------------------------------------------------------
// Driver analysis
// ---------------------------------------------------------------------------

/**
 * Result of analyzing lifecycle steps to determine what prerequisites
 * (auth, Azure subscription, etc.) are needed before execution.
 */
export interface LifecycleAnalysis {
  /** Whether any step uses an M365-scoped driver. */
  needsM365: boolean;
  /** Whether any step uses an Azure-scoped driver. */
  needsAzure: boolean;
  /** All driver IDs referenced in the steps. */
  driverIds: string[];
  /** Unresolved ${{VAR}} placeholders found in step configs. */
  unresolvedVars: string[];
}

// ---------------------------------------------------------------------------
// Prerequisite result types
// ---------------------------------------------------------------------------

/**
 * M365 tenant information obtained after authentication.
 */
export interface M365TenantInfo {
  tenantId: string;
  displayName?: string;
}

/**
 * Azure account information obtained after authentication.
 */
export interface AzureAccountInfo {
  accountId?: string;
  tenantId?: string;
}

/**
 * Re-export SubscriptionInfo from the API layer.
 */
export type { SubscriptionInfo } from "../api/utils/login";

/**
 * Information about the target Azure resource group.
 */
export interface ResourceGroupInfo {
  name: string;
  location: string;
  isNew: boolean;
}

// ---------------------------------------------------------------------------
// Lifecycle operation result
// ---------------------------------------------------------------------------

/**
 * Post-execution action the consumer should render (URL link, message, etc.).
 */
export interface PostAction {
  type: "openUrl" | "showMessage";
  message: string;
  url?: string;
}

/**
 * Enriched result returned by lifecycle operations (provision, deploy, publish).
 */
export interface LifecycleOperationResult {
  /** Core execution result from the lifecycle engine. */
  lifecycleResult: LifecycleResult;
  /** Environment variables after execution (includes driver outputs). */
  envMap: Map<string, string>;
  /** Suggested post-execution actions for the consumer to render. */
  postActions: PostAction[];
}
