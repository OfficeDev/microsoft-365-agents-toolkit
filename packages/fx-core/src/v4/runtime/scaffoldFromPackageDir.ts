// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result, err } from "neverthrow";
import { Answers, CallerFloor } from "../model/dataModel";
import { ScaffoldOutcome, TargetDir } from "../pipeline/runScaffoldPipeline";
import { loadPackageDir } from "../distribution/packageDir";
import { createRealRuntime } from "./realRuntime";
import { scaffold } from "./scaffold";

/** Product front door for scaffolding an on-disk declarative package. */
export async function scaffoldFromPackageDir(
  packageDir: string,
  answers: Answers,
  callerFloor: CallerFloor,
  targetDir: TargetDir,
  flagReader?: (name: string) => boolean
): Promise<Result<ScaffoldOutcome, FxError>> {
  const loaded = loadPackageDir(packageDir);
  if (loaded.isErr()) {
    return err(loaded.error);
  }
  const runtime = createRealRuntime(targetDir.path, flagReader);
  return scaffold(
    {
      descriptor: loaded.value.descriptor,
      pipeline: loaded.value.pipeline,
      content: loaded.value.content,
      answers,
      callerFloor,
      targetDir,
    },
    runtime
  );
}
