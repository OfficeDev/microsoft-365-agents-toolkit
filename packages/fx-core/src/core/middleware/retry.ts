// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, Result, SystemError } from "@microsoft/teamsfx-api";
import { assembleError } from "../../error";
import { TOOLS } from "../../common/globalVars";

export async function runWithRetry<T>(
  run: () => Promise<Result<T, FxError>>,
  retryCondition: (result: Result<T, FxError>, attempt: number) => boolean,
  maxRetry = 3
): Promise<Result<T, FxError>> {
  let lastResult: Result<T, FxError> = err(
    new SystemError(
      "RetryMiddleware",
      "UninitializedResult",
      "runWithRetry did not execute the run function."
    )
  );
  for (let i = 0; i < maxRetry; i++) {
    try {
      lastResult = await run();
      if (!retryCondition(lastResult, i + 1)) {
        return lastResult;
      }
    } catch (e) {
      const error = assembleError(e);
      if (!retryCondition(err(error), i + 1)) {
        throw e;
      }
    }
    TOOLS?.logProvider?.info(`Retrying operation ... (${i + 1}/${maxRetry})`);
  }
  return lastResult;
}
