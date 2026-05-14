// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError, LifecycleOperationResult } from "@microsoft/teamsfx-core-next";
import { runOperation, provisionOp, deployOp, publishOp } from "@microsoft/teamsfx-core-next";

export interface LifecycleInput {
  projectPath: string;
  envName: string;
  skipConsent?: boolean;
}

export interface PublishInput {
  projectPath: string;
  envName: string;
}

export interface LifecycleOutput {
  result: LifecycleOperationResult;
  postActions: LifecycleOperationResult["postActions"];
}

export async function provisionAction(
  ctx: AtkContext,
  input: LifecycleInput
): Promise<LifecycleOutput> {
  const result = await runOperation(provisionOp, ctx, {
    projectPath: input.projectPath,
    envName: input.envName,
    skipConsent: input.skipConsent,
  });
  if (result.isErr()) throw toError(result.error);
  return { result: result.value, postActions: result.value.postActions };
}

export async function deployAction(
  ctx: AtkContext,
  input: LifecycleInput
): Promise<LifecycleOutput> {
  const result = await runOperation(deployOp, ctx, {
    projectPath: input.projectPath,
    envName: input.envName,
    skipConsent: input.skipConsent,
  });
  if (result.isErr()) throw toError(result.error);
  return { result: result.value, postActions: result.value.postActions };
}

export async function publishAction(
  ctx: AtkContext,
  input: PublishInput
): Promise<LifecycleOutput> {
  const result = await runOperation(publishOp, ctx, {
    projectPath: input.projectPath,
    envName: input.envName,
  });
  if (result.isErr()) throw toError(result.error);
  return { result: result.value, postActions: result.value.postActions };
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
