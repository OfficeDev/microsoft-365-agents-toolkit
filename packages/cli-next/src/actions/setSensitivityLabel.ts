// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { runOperation, declarativeAgent } from "@microsoft/teamsfx-core-next";

export interface SetSensitivityLabelInput {
  projectPath: string;
  agentManifestPath?: string;
  labelId: string;
}

export async function setSensitivityLabelAction(
  ctx: AtkContext,
  input: SetSensitivityLabelInput
): Promise<void> {
  const agentManifestPath = await resolveAgentManifestPath(
    input.agentManifestPath,
    input.projectPath
  );

  const result = await runOperation(declarativeAgent.setSensitivityLabelOp, ctx, {
    agentManifestPath,
    labelId: input.labelId,
  });
  if (result.isErr()) throw toError(result.error);
}

async function resolveAgentManifestPath(
  explicit: string | undefined,
  projectPath: string
): Promise<string> {
  if (explicit) return explicit;
  const result = await declarativeAgent.getAgentManifestPath(projectPath);
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
