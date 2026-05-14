// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { runOperation, teamsApp } from "@microsoft/teamsfx-core-next";

export interface ValidateInput {
  manifestPath?: string;
  packagePath?: string;
}

export interface ValidateOutput {
  valid: boolean;
  errors: string[];
  warnings?: string[];
}

export async function validateAction(
  ctx: AtkContext,
  input: ValidateInput
): Promise<ValidateOutput> {
  if (input.packagePath) {
    const result = await runOperation(teamsApp.validateAppPackageOp, ctx, {
      packagePath: input.packagePath,
    });
    if (result.isErr()) throw toError(result.error);
    return { valid: result.value.valid, errors: result.value.errors };
  }

  if (input.manifestPath) {
    const result = await runOperation(teamsApp.validateManifestOp, ctx, {
      manifestPath: input.manifestPath,
    });
    if (result.isErr()) throw toError(result.error);
    return {
      valid: result.value.valid,
      errors: result.value.errors,
      warnings: result.value.warnings,
    };
  }

  throw new Error("Either --manifest-path or --app-package-file-path is required.");
}

export interface PackageInput {
  projectPath: string;
  manifestPath?: string;
  outputPath: string;
}

export interface PackageOutput {
  zipPath: string;
}

export async function packageAction(ctx: AtkContext, input: PackageInput): Promise<PackageOutput> {
  const result = await runOperation(teamsApp.packageAppOp, ctx, {
    projectPath: input.projectPath,
    manifestPath: input.manifestPath,
    outputPath: input.outputPath,
  });
  if (result.isErr()) throw toError(result.error);
  return { zipPath: result.value.zipPath ?? input.outputPath };
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
