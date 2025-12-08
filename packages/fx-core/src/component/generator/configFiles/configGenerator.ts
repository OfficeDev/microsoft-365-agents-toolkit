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
    components: { name: string; programmingLanguage: string }[]
  ): Promise<Result<GeneratorResult, FxError>> {
    await context.userInteraction.showMessage("info", "Generating configuration files...", false);
    for (const component of components) {
      const policy = policys[`${component.name}-${component.programmingLanguage}`];
      const fileDetectionResult = await this.detectFileConflict(destinationPath, policy);
      if (fileDetectionResult.isErr()) {
        return err(fileDetectionResult.error);
      }
    }
    for (const component of components) {
      const policy = policys[`${component.name}-${component.programmingLanguage}`];
      const sourcePath = path.join(
        getTemplatesFolder(),
        "configs",
        component.name,
        component.programmingLanguage
      );
      await this.generateConfigFilesByPolicy(sourcePath, destinationPath, policy);
    }
    return ok({});
  }

  private async detectFileConflict(
    destinationPath: string,
    policy: Record<string, CopyPolicy>
  ): Promise<Result<void, FxError>> {
    for (const [filePath, copyPolicy] of Object.entries(policy)) {
      // Here we should check if the file exists in the destinationPath.
      const fullPath = path.join(destinationPath, filePath);
      const fileExists = await fs.pathExists(fullPath);
      if (fileExists) {
        if (copyPolicy.allowExistingFile === false) {
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

  private async generateConfigFilesByPolicy(
    sourcePath: string,
    destinationPath: string,
    policy: Record<string, CopyPolicy>
  ): Promise<void> {
    for (const [filePath, copyPolicy] of Object.entries(policy)) {
      const srcFilePath = path.join(sourcePath, filePath);
      const destFilePath = path.join(destinationPath, filePath);
      const fileExists = await fs.pathExists(destFilePath);
      if (fileExists) {
        if (copyPolicy.policy === "add") {
          if (srcFilePath.endsWith(".json")) {
            await mergeJsonFile(srcFilePath, destFilePath);
          }
        }
      } else {
        // If the file does not exist, just copy it.
        await fs.copy(srcFilePath, destFilePath);
      }
    }
  }
}

export const configGenerator = new ConfigGenerator();
