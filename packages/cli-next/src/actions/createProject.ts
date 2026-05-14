// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core-next";
import { project } from "@microsoft/teamsfx-core-next";
import { colorize, TextType } from "../output";

export interface CreateProjectInput {
  templateId: string;
  projectName: string;
  language: string;
  destinationPath: string;
  options?: Record<string, unknown>;
}

export interface CreateProjectOutput {
  projectPath: string;
  projectId: string;
  warnings?: string[];
}

/**
 * Create a new project from a template.
 *
 * In non-interactive mode, calls runOperation(createProjectOp) directly.
 * In interactive mode (future), pre-fills inputs and delegates to createProjectInteractive.
 */
export async function createProjectAction(
  ctx: AtkContext,
  input: CreateProjectInput
): Promise<CreateProjectOutput> {
  const { runOperation } = await import("@microsoft/teamsfx-core-next");

  const result = await runOperation(project.createProjectOp, ctx, {
    templateId: input.templateId,
    projectName: input.projectName,
    language: input.language,
    destinationPath: input.destinationPath,
    options: input.options,
  });

  if (result.isErr()) {
    throw toError(result.error);
  }

  const value = result.value;
  console.log(colorize(`Project created at: ${value.projectPath}`, TextType.Success));
  if (value.warnings?.length) {
    for (const w of value.warnings) {
      console.warn(colorize(`Warning: ${w}`, TextType.Warning));
    }
  }

  return {
    projectPath: value.projectPath,
    projectId: value.projectId,
    warnings: value.warnings,
  };
}

/**
 * Convert an AtkError to a throwable Error (caught by wrapHandler).
 */
function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
