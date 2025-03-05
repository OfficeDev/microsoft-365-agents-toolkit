// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import { FxError, LogProvider, Result } from "@microsoft/teamsfx-api";
import { DriverContext } from "../driver/interface/commonArgs";
import { StepDriver } from "../driver/interface/stepDriver";

export type AdditionalMetadata = {
  [key: string]: unknown;
};

export type RawProjectModel = {
  registerApp?: DriverDefinition[];
  provision?: DriverDefinition[];
  configureApp?: DriverDefinition[];
  deploy?: DriverDefinition[];
  publish?: DriverDefinition[];
  share?: DriverDefinition[];
  environmentFolderPath?: string;
  version: string;
  additionalMetadata?: AdditionalMetadata;
};

export type ProjectModel = {
  registerApp?: ILifecycle;
  provision?: ILifecycle;
  configureApp?: ILifecycle;
  deploy?: ILifecycle;
  publish?: ILifecycle;
  share?: ILifecycle;
  environmentFolderPath?: string;
  version: string;
  additionalMetadata?: AdditionalMetadata;
};

export type DriverDefinition = {
  name?: string;
  uses: string;
  with: unknown;
  env?: Record<string, string>;
  writeToEnvironmentFile?: Record<string, string>;
};

export type DriverInstance = DriverDefinition & { instance: StepDriver };

export type LifecycleNames = [
  "registerApp",
  "configureApp",
  "provision",
  "deploy",
  "publish",
  "share"
];
export const LifecycleNames: LifecycleNames = [
  "registerApp",
  "configureApp",
  "provision",
  "deploy",
  "publish",
  "share",
];
type AnyElementOf<T extends unknown[]> = T[number];
export type LifecycleName = AnyElementOf<LifecycleNames>;

export type UnresolvedPlaceholders = string[];
export type ResolvedPlaceholders = string[];

export type Output = { env: Map<string, string>; unresolvedPlaceHolders: UnresolvedPlaceholders };

export type PartialSuccessReason =
  | { kind: "DriverError"; failedDriver: DriverDefinition; error: FxError }
  | {
      kind: "UnresolvedPlaceholders";
      failedDriver: DriverDefinition;
      unresolvedPlaceHolders: UnresolvedPlaceholders;
    };

export type ExecutionOutput = Map<string, string>;

export type ExecutionError =
  | { kind: "PartialSuccess"; env: Map<string, string>; reason: PartialSuccessReason }
  | { kind: "Failure"; error: FxError };

export type ExecutionResult = {
  result: Result<ExecutionOutput, ExecutionError>;
  summaries: string[][];
};

export interface ILifecycle {
  name: LifecycleName;
  driverDefs: DriverDefinition[];

  /**
   * Resolve all placeholders in the driver's arguments based on the environment variables in-place.
   * Unresolved placeholders will be returned. It can be used to get unresolved placeholders before actually
   * executing a lifecycle. Useful for getting unresolved built-in placeholders like AZURE_SUBSCRIPTION_ID
   * and RESOURCE_GROUP and asking for user input.
   * @returns unresolved placeholder names
   */
  resolvePlaceholders(): UnresolvedPlaceholders;

  /**
   * execute() will run drivers one by one. The difference between execute() and run()
   * is: 1. execute() resolves a driver's placeholder before executing it. It's useful when driver2 references
   *      driver1's output.
   *     2. execute() still returns the output of successful driver runs when encountering an error.
   *     3. execute() returns a list of summaires
   * @param ctx driver context
   */
  execute(ctx: DriverContext): Promise<ExecutionResult>;

  /**
   * Try to search for driver instances defined by this.driverDefs.
   * @param log LogProvider
   */
  resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError>;
}

export interface IYamlParser {
  parse(path: string, validateSchema?: boolean): Promise<Result<ProjectModel, FxError>>;
}
