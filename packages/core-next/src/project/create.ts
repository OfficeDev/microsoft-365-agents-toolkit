// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import * as crypto from "crypto";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { defineOperation, runOperation } from "../core/operation";
import type { AtkContext } from "../core/context";
import { userError, systemError } from "../core/error";
import { templateRegistry } from "../templates/registry";
import type { TemplateActionOptions, Language } from "../templates/types";
import { buildQuestionTree } from "../questions/treeBuilder";
import { traverseQuestionTree } from "../questions/traverse";
import { QuestionNames } from "../questions/questionNames";
import type { Inputs } from "../api/types";

/**
 * Input schema for createProject operation.
 */
const createProjectInputSchema = z.object({
  /** Template ID from the TemplateRegistry */
  templateId: z.string().min(1),
  /** Human-readable project name */
  projectName: z.string().min(1).max(128),
  /** Programming language selection */
  language: z.enum(["typescript", "javascript", "python", "csharp", "common"]),
  /** Parent directory where the project folder will be created */
  destinationPath: z.string().min(1),
  /** Additional template-specific options */
  options: z.record(z.unknown()).optional(),
});

type CreateProjectInput = z.infer<typeof createProjectInputSchema>;

/**
 * Result of a successful project creation.
 */
interface CreateProjectResult {
  projectPath: string;
  projectId: string;
  warnings?: string[];
}

/**
 * Create a new project from a registered template.
 *
 * Pipeline: validate input → lookup template → scaffold → ensure tracking ID
 */
export const createProjectOp = defineOperation(
  "createProject",
  createProjectInputSchema,
  async (ctx: AtkContext, input: CreateProjectInput) => {
    const { templateId, projectName, language, destinationPath, options } = input;

    // 1. Lookup template
    const descriptor = templateRegistry.get(templateId);
    if (!descriptor) {
      return err(
        userError("TemplateNotFound", `No template registered with ID "${templateId}".`, {
          source: "project/create",
        })
      );
    }

    // 2. Verify language support
    if (!descriptor.languages.includes(language as Language)) {
      return err(
        userError(
          "UnsupportedLanguage",
          `Template "${templateId}" does not support language "${language}". ` +
            `Supported: ${descriptor.languages.join(", ")}`,
          { source: "project/create" }
        )
      );
    }

    // 3. Create destination folder
    const projectPath = path.join(destinationPath, projectName);
    try {
      await fs.promises.mkdir(projectPath, { recursive: true });
    } catch (e) {
      return err(
        systemError("CreateDirectoryFailed", `Failed to create project directory: ${e}`, {
          source: "project/create",
          inner: e instanceof Error ? e : new Error(String(e)),
        })
      );
    }

    // 4. Scaffold
    const opts: TemplateActionOptions = {
      language: language as Language,
      projectName,
      destinationPath: projectPath,
      ...(options ?? {}),
    };

    const scaffoldResult = await descriptor.scaffoldFn(ctx, opts);
    if (scaffoldResult.isErr()) {
      return err(scaffoldResult.error);
    }

    // 5. Ensure tracking ID
    const projectId = await ensureTrackingId(projectPath);

    return ok({
      projectPath,
      projectId,
      warnings: scaffoldResult.value.warnings,
    } satisfies CreateProjectResult);
  }
);

/**
 * Create a new project by walking the full question tree interactively.
 *
 * Use this when the caller has partial or no inputs (e.g. the VS Code "Create"
 * command or CLI with missing args). The question tree is built from the
 * template registry, traversed via the platform's UserInteraction, and then
 * the result is handed to `createProjectOp`.
 */
export async function createProjectInteractive(
  ctx: AtkContext,
  inputs: Inputs
): Promise<
  | { ok: true; value: CreateProjectResult }
  | { ok: false; error: ReturnType<typeof userError> | ReturnType<typeof systemError> }
> {
  // Build & traverse the question tree
  const tree = buildQuestionTree(templateRegistry);
  const traverseResult = await traverseQuestionTree(tree, ctx.ui, inputs);
  if (traverseResult.isErr()) {
    return { ok: false, error: traverseResult.error };
  }

  // Extract answers collected by traversal
  const templateId = inputs[QuestionNames.templateId] as string | undefined;
  const projectName = inputs[QuestionNames.projectName] as string | undefined;
  const language = inputs[QuestionNames.language] as string | undefined;
  const folder = inputs[QuestionNames.destinationFolder] as string | undefined;

  if (!templateId || !projectName || !language || !folder) {
    return {
      ok: false,
      error: userError(
        "IncompleteInputs",
        "Question traversal did not collect all required inputs.",
        {
          source: "project/create",
        }
      ),
    };
  }

  // Delegate to the validated createProjectOp
  const result = await runOperation(createProjectOp, ctx, {
    templateId,
    projectName,
    language,
    destinationPath: folder,
    options: inputs,
  });

  if (result.isErr()) {
    return { ok: false, error: result.error };
  }

  return { ok: true, value: result.value };
}

/**
 * Ensure the project has a unique tracking ID in the YAML config.
 * If one doesn't exist, generates a UUID and writes it.
 */
async function ensureTrackingId(projectPath: string): Promise<string> {
  const yamlPath = path.join(projectPath, "teamsapp.yaml");
  const id = crypto.randomUUID();
  try {
    let content = await fs.promises.readFile(yamlPath, "utf-8");
    if (!content.includes("projectId:")) {
      content = `projectId: ${id}\n${content}`;
      await fs.promises.writeFile(yamlPath, content, "utf-8");
    }
  } catch {
    // YAML file may not exist yet for some templates — that's ok
  }
  return id;
}
