// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { runOperation, declarativeAgent } from "@microsoft/teamsfx-core-next";

export interface AddAuthConfigInput {
  projectPath: string;
  authType: "oauth" | "api-key";
  ymlPath?: string;
  authName: string;
  specPath: string;
  entra?: boolean;
  enablePkce?: boolean;
  registrationId?: string;
}

export async function addAuthConfigAction(
  ctx: AtkContext,
  input: AddAuthConfigInput
): Promise<void> {
  const ymlPath = input.ymlPath ?? `${input.projectPath}/teamsapp.yml`;

  if (input.authType === "oauth") {
    const result = await runOperation(declarativeAgent.injectOAuthActionOp, ctx, {
      ymlPath,
      authName: input.authName,
      specRelativePath: input.specPath,
      isMicrosoftEntra: input.entra ?? false,
      enablePKCE: input.enablePkce,
      registrationId: input.registrationId,
    });
    if (result.isErr()) throw toError(result.error);
  } else {
    const result = await runOperation(declarativeAgent.injectApiKeyActionOp, ctx, {
      ymlPath,
      authName: input.authName,
      specRelativePath: input.specPath,
      registrationId: input.registrationId,
    });
    if (result.isErr()) throw toError(result.error);
  }
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
