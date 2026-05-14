// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { environment, type AtkError } from "@microsoft/teamsfx-core-next";

export async function envListAction(projectPath: string): Promise<string[]> {
  const result = await environment.listEnvironments(projectPath);
  if (result.isErr()) throw toError(result.error);
  return result.value;
}

export async function envAddAction(
  projectPath: string,
  name: string,
  copyFrom?: string
): Promise<void> {
  const result = await environment.addEnvironment(projectPath, name, copyFrom);
  if (result.isErr()) throw toError(result.error);
}

export async function envResetAction(projectPath: string, name: string): Promise<void> {
  const result = await environment.resetEnvironment(projectPath, name);
  if (result.isErr()) throw toError(result.error);
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
