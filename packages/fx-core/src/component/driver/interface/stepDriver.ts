// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DriverContext } from "./commonArgs";
import { FxError, Result } from "@microsoft/teamsfx-api";

export type ExecutionResult = {
  result: Result<Map<string, string>, FxError>;
  // summaries hold the successful logs even if the result is a FxError.
  summaries: string[];
  // Preserve domain receipts after a late cancellation; never perform recovery here.
  onCancellation?: () => FxError;
};

export interface StepDriver {
  readonly description?: string;
  readonly progressTitle?: string;

  /**
   * Run the driver and return success summary entries along with the result, no matter the result is success or failure.
   * Because a failed action may still emit some succuessful summaires.
   * The summary is expected to contain human readable information that will be presented to users.
   * @param args Arguments from the `with` section in the yaml file.
   * @param ctx logger, telemetry, progress bar, etc.
   * @param outputEnvVarNames the environment variable names for each output
   * @param schemaVersion schema version of the executed yaml file
   * @param name name of the executed yaml section
   */
  execute(
    args: unknown,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>,
    schemaVersion?: string,
    name?: string
  ): Promise<ExecutionResult>;
}
