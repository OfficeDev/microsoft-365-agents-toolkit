// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import { Context } from "./context";
import { FxError } from "./error";
import { Inputs, Warning } from "./types";

export interface IGenerator {
  componentName: string;
  run(
    context: Context,
    inputs: Inputs,
    destinationPath: string
  ): Promise<Result<GeneratorResult, FxError>>;
}

export interface GeneratorResult {
  warnings?: Warning[];
}
