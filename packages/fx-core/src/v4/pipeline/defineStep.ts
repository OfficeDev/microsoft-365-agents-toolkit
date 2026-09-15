// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import type { PreparedStep, RegisteredStep, StepContext, StepParams } from "./runScaffoldPipeline";

export function defineStep<Params>(definition: {
  parse(resolved: StepParams): Result<Params, string>;
  apply(params: Params, ctx: StepContext): Result<void, FxError> | Promise<Result<void, FxError>>;
  invalidParams(violation: string): FxError;
}): RegisteredStep {
  const prepare = (resolved: StepParams): Result<PreparedStep, string> =>
    definition.parse(resolved).map((params) => (ctx: StepContext) => definition.apply(params, ctx));
  return {
    prepare,
    validateParams(resolved) {
      const parsed = definition.parse(resolved);
      return parsed.isErr() ? parsed.error : undefined;
    },
    apply(resolved, ctx) {
      const prepared = prepare(resolved);
      return prepared.isErr() ? err(definition.invalidParams(prepared.error)) : prepared.value(ctx);
    },
  };
}

export function prepareStep(
  step: RegisteredStep,
  resolved: StepParams
): Result<PreparedStep, string> {
  if (step.prepare !== undefined) return step.prepare(resolved);
  const violation = step.validateParams(resolved);
  return violation === undefined ? ok((ctx) => step.apply(resolved, ctx)) : err(violation);
}
