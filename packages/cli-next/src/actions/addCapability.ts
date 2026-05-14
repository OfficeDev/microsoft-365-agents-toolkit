// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { runOperation, declarativeAgent } from "@microsoft/teamsfx-core-next";

export interface AddCapabilityInput {
  projectPath: string;
  agentManifestPath?: string;
  source: "web-search" | "onedrive-sharepoint" | "graph-connector" | "embedded-knowledge";
  siteUrl?: string;
  connectionIds?: string[];
  filePaths?: string[];
}

export async function addCapabilityAction(
  ctx: AtkContext,
  input: AddCapabilityInput
): Promise<void> {
  const agentManifestPath = await resolveAgentManifestPath(
    input.agentManifestPath,
    input.projectPath
  );

  const result = await runOperation(declarativeAgent.addKnowledgeOp, ctx, {
    agentManifestPath,
    source: input.source,
    siteUrl: input.siteUrl,
    connectionIds: input.connectionIds,
    embeddedFilePaths: input.filePaths,
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
