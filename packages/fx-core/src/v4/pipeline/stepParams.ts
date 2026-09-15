// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { StepParams } from "./runScaffoldPipeline";

export function stringParam(params: StepParams, key: string): string | undefined {
  const value = params[key];
  return typeof value === "string" ? value : undefined;
}

export function stringArrayParam(params: StepParams, key: string): string[] | undefined {
  const value = params[key];
  return Array.isArray(value) && value.every((item) => typeof item === "string")
    ? value
    : undefined;
}
