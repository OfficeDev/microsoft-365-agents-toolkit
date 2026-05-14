// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { runOperation, teamsApp } from "@microsoft/teamsfx-core-next";

export interface M365SideloadInput {
  filePath: string;
  scope?: string;
}

export async function m365SideloadAction(
  ctx: AtkContext,
  input: M365SideloadInput
): Promise<{ titleId: string; appId: string; shareLink: string }> {
  const result = await runOperation(teamsApp.extendToM365Op, ctx, {
    appPackagePath: input.filePath,
    scope: input.scope,
  });
  if (result.isErr()) throw toError(result.error);
  return result.value;
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
