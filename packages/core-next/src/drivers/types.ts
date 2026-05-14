// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import { AtkContext } from "../core/context";
import { AtkError } from "../core/error";

/**
 * Output of a driver execution.
 */
export interface DriverOutput {
  /** Key-value pairs produced by the driver (e.g. resource IDs, endpoints) */
  outputs: Record<string, string>;
  /** Warnings emitted during execution */
  warnings?: string[];
}

/**
 * Configuration passed to a driver, typically from YAML lifecycle entries.
 */
export interface DriverConfig {
  /** Driver-specific configuration from m365agents.yml */
  [key: string]: unknown;
}

/**
 * DriverDescriptor is a plain data object describing a service-interaction driver.
 * Replaces the Driver class hierarchy with a flat, registrable descriptor.
 *
 * Drivers are self-contained units that perform one atomic operation against
 * a service (Azure, Teams Platform, AAD/Entra, etc.).
 */
export interface DriverDescriptor {
  /** Unique driver ID matching YAML references (e.g. "azure-app-service-deploy") */
  id: string;

  /** Human-readable display name */
  name: string;

  /** Execute the driver's primary action */
  executeFn: (ctx: AtkContext, config: DriverConfig) => Promise<Result<DriverOutput, AtkError>>;

  /** Optional rollback function for undoing the action */
  rollbackFn?: (ctx: AtkContext, config: DriverConfig) => Promise<Result<void, AtkError>>;

  /** Optional validation function for config before execution */
  validateFn?: (config: DriverConfig) => Result<void, AtkError>;
}
