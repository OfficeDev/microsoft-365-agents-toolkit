// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  Context,
  err,
  FxError,
  GeneratorResult,
  ok,
  Result,
  UserError,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { getDefaultString } from "../../../common/localizeUtils";
import { TelemetryEvent } from "../../../common/telemetry";
import { getTemplatesFolder } from "../../../folder";
import { ProgressTitles } from "../../messages";
import { ActionExecutionMW } from "../../middleware/actionExecutionMW";
import { CopyPolicy, policys } from "./copyPolicy";
import { mergeJsonFile } from "./jsonMerger";
import { renderTemplate } from "./renderTemplate";

export class ConfigGenerator {
  componentName = "ConfigGenerator";

  @hooks([
    ActionExecutionMW({
      enableProgressBar: true,
      progressTitle: ProgressTitles.create,
      progressSteps: 1,
      enableTelemetry: true,
      telemetryEventName: TelemetryEvent.GenerateConfig,
    }),
  ])
  public async run(
    context: Context,
    destinationPath: string,
    components: { name: string; programmingLanguage: string }[],
    features: Record<string, unknown>
  ): Promise<Result<GeneratorResult, FxError>> {
    await context.userInteraction.showMessage("info", "Generating configuration files...", false);

    // Process all components: detect conflicts and generate files in a single pass
    for (const component of components) {
      const policyKey = this.getPolicyKey(component);
      const policy = policys[policyKey];

      if (!policy) {
        return err(
          new UserError(
            this.componentName,
            "UnknownPolicyError",
            getDefaultString("error.generator.UnknownPolicy", policyKey)
          )
        );
      }

      const fileDetectionResult = await this.detectFileConflict(destinationPath, policy);
      if (fileDetectionResult.isErr()) {
        await context.userInteraction.showMessage("warn", fileDetectionResult.error.message, false);
        continue;
      }

      const sourcePath = path.join(
        getTemplatesFolder(),
        "configs",
        component.name,
        component.programmingLanguage
      );
      await this.generateConfigFilesByPolicy(sourcePath, destinationPath, policy, features);
    }
    return ok({});
  }

  private getPolicyKey(component: { name: string; programmingLanguage: string }): string {
    return `${component.name}-${component.programmingLanguage}`;
  }

  private async detectFileConflict(
    destinationPath: string,
    policy: Record<string, CopyPolicy>
  ): Promise<Result<void, FxError>> {
    for (const [filePath, copyPolicy] of Object.entries(policy)) {
      if (!copyPolicy.allowExistingFile) {
        const fullPath = path.join(destinationPath, filePath);
        const fileExists = await fs.pathExists(fullPath);
        if (fileExists) {
          return err(
            new UserError(
              this.componentName,
              "ConflictFileError",
              getDefaultString("error.generator.FileConflictError", filePath)
            )
          );
        }
      }
    }
    return ok(undefined);
  }

  private getFileExtensionWithoutTemplate(filePath: string): string {
    const withoutTemplate = filePath.endsWith(".tpl") ? filePath.slice(0, -4) : filePath;
    return path.extname(withoutTemplate);
  }

  private async generateConfigFilesByPolicy(
    sourcePath: string,
    destinationPath: string,
    policy: Record<string, CopyPolicy>,
    features: Record<string, unknown>
  ): Promise<void> {
    for (const [filePath, copyPolicy] of Object.entries(policy)) {
      const isTemplate = filePath.endsWith(".tpl");
      let srcFilePath = path.join(sourcePath, filePath);
      const destFilePath = path.join(
        destinationPath,
        isTemplate ? filePath.slice(0, -4) : filePath
      );
      const fileExtension = this.getFileExtensionWithoutTemplate(filePath);
      let renderedFilePath: string | null = null;

      // Render template if needed
      if (isTemplate) {
        renderedFilePath = destFilePath + ".rendered";
        const renderedContent = renderTemplate(srcFilePath, features);
        await fs.writeFile(renderedFilePath, renderedContent, "utf-8");
        srcFilePath = renderedFilePath;
      }

      try {
        // Handle existing files
        const fileExists = await fs.pathExists(destFilePath);
        if (fileExists) {
          if (copyPolicy.policy === "add" && fileExtension === ".json") {
            await mergeJsonFile(srcFilePath, destFilePath);
          }
          // For "skip" or non-JSON files, do nothing
        } else {
          // If the file does not exist, just copy it.
          await fs.copy(srcFilePath, destFilePath);
        }
      } finally {
        // Clean up rendered temp file
        if (renderedFilePath !== null) {
          await fs.remove(renderedFilePath);
        }
      }
    }
  }
}

export const configGenerator = new ConfigGenerator();
