// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { ok, err, Result } from "neverthrow";
import type { AtkContext } from "../../core/context";
import type { AtkError } from "../../core/error";
import { userError } from "../../core/error";
import type { TemplateActionOptions, ScaffoldResult } from "../types";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import type { TemplateInfo } from "../scaffold/types";
import type { SpecParserAdapter } from "./specParserAdapter";
import { createSpecParserAdapter } from "./specParserAdapter";

/**
 * Project type determines which base template and post-processing to use.
 */
export type OpenApiProjectType = "Copilot" | "TeamsAi" | "SME";

/**
 * Create a scaffold function for OpenAPI-spec-driven templates.
 *
 * Pipeline:
 * 1. Scaffold the base template (download + Mustache render)
 * 2. Validate the API spec via SpecParserAdapter
 * 3. Generate code from selected operations
 * 4. Write generated files into the scaffolded project
 *
 * @param projectType - Copilot (DA), TeamsAi (AI Agent), or SME (Message Extension)
 * @param baseTemplateName - The artifact name for the base template ZIP
 * @param adapter - Optional SpecParserAdapter override (defaults to stub)
 */
export function makeOpenApiScaffoldFn(
  projectType: OpenApiProjectType,
  baseTemplateName: string,
  adapter?: SpecParserAdapter
) {
  const specParser = adapter ?? createSpecParserAdapter();

  return async (
    ctx: AtkContext,
    opts: TemplateActionOptions
  ): Promise<Result<ScaffoldResult, AtkError>> => {
    const apiSpecPath = opts.apiSpecPath as string | undefined;
    const apiOperations = opts.apiOperations as string[] | undefined;

    if (!apiSpecPath) {
      return err(
        userError("MissingApiSpecPath", "API specification path is required.", {
          source: `openApi/${projectType}`,
        })
      );
    }

    if (!apiOperations || apiOperations.length === 0) {
      return err(
        userError("MissingApiOperations", "At least one API operation must be selected.", {
          source: `openApi/${projectType}`,
        })
      );
    }

    // Step 1: Scaffold the base template
    const replaceMap: Record<string, string> = {
      ...getTemplateReplaceMap({
        appName: opts.projectName,
        ...opts,
      }),
    };

    if (projectType === "Copilot") {
      replaceMap.DeclarativeCopilot = "true";
    }

    const tplInfo: TemplateInfo = {
      templateName: baseTemplateName,
      language: opts.language === "common" ? "common" : opts.language.substring(0, 2),
      replaceMap,
    };

    const scaffoldResult = await scaffoldTemplates(ctx, [tplInfo], opts.destinationPath);
    if (scaffoldResult.isErr()) {
      return err(scaffoldResult.error);
    }

    // Step 2: Validate the API spec
    const validation = await specParser.validate(apiSpecPath);
    if (!validation.valid) {
      return err(
        userError(
          "InvalidApiSpec",
          `API specification validation failed: ${validation.errors.join("; ")}`,
          { source: `openApi/${projectType}` }
        )
      );
    }

    // Step 3: Generate code from operations
    const genResult = await specParser.generate(
      apiSpecPath,
      apiOperations,
      opts.destinationPath,
      projectType
    );

    // Step 4: Write generated files
    for (const [filePath, content] of genResult.files) {
      const fullPath = path.join(opts.destinationPath, filePath);
      await fs.promises.mkdir(path.dirname(fullPath), { recursive: true });
      await fs.promises.writeFile(fullPath, content, "utf-8");
    }

    const warnings = [
      ...(scaffoldResult.value.length === 0 ? ["No base template files were scaffolded"] : []),
      ...genResult.warnings,
    ];

    return ok({
      projectPath: opts.destinationPath,
      warnings: warnings.length > 0 ? warnings : undefined,
    });
  };
}
