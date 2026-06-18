// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result, err } from "neverthrow";
import { Answers, CallerFloor } from "../model/dataModel";
import { ScaffoldOutcome, TargetDir } from "../pipeline/runScaffoldPipeline";
import { loadPackageDir } from "../distribution/packageDir";
import { createRealRuntime } from "./realRuntime";
import { scaffold } from "./scaffold";

/**
 * The product front-door for the on-disk declarative path: load an authored
 * template package from `packageDir`, then scaffold it onto `targetDir` through
 * the on-disk runtime. This is the single call a surface (or the v3 bridge)
 * makes to run a declarative create end-to-end — the engine composition the T3
 * scenario proves in memory, here driven entirely by product code rather than
 * test-side request assembly.
 *
 * It adds no behavior of its own: `loadPackageDir` owns package I/O,
 * `createRealRuntime` owns the on-disk file sink, and `scaffold` owns the
 * parse → render-context → two-phase-pipeline weld. Every fact a caller asserts
 * is owned by one of those composed pieces.
 *
 * Welding this onto the live distribution chain (so the bundled declarative
 * package reaches it) is the follow-on, blocked on distributing the declarative
 * descriptor/pipeline format — the v4 channel currently ships v3-shaped zips.
 *
 * v4-owned (INV-7): imports no v3 symbol.
 */
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
