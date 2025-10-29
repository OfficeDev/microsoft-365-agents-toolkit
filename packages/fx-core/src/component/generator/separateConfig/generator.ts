// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Context,
  err,
  FxError,
  GeneratorResult,
  Inputs,
  ok,
  Result,
  SystemError,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { getTemplatesFolder } from "../../../folder";
import { ProgrammingLanguage, QuestionNames } from "../../../question/constants";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { getAllTemplatesOnPlatform } from "../templates/metadata";
import { scaffoldConfigComponents } from "./configComponent";
import { TemplateConfig } from "./interface";

export class SeparateConfigGenerator extends DefaultTemplateGenerator {
  // override this property to send telemetry event with different component name
  componentName = "separate-config-generator";

  public override activate(context: Context, inputs: Inputs): boolean {
    const templateName = inputs[QuestionNames.TemplateName];
    return ["non-sso-tab", "custom-copilot-rag-azure-ai-search"].includes(templateName);
  }

  protected override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const templateName = inputs[QuestionNames.TemplateName] as string;
    const language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    const templateMetadata = getAllTemplatesOnPlatform(inputs.platform).find(
      (t) => t.name === templateName && t.language === language
    );
    if (!templateMetadata) {
      return Promise.resolve(
        err(
          new SystemError(
            "SeparateConfigGenerator",
            "TemplateNotFound",
            `Template ${templateName} with language ${language} not found`
          )
        )
      );
    }
    const templateConfigMetadata = await this.getTemplateConfigMetadata(templateMetadata.id);
    await scaffoldConfigComponents(templateConfigMetadata, destinationPath, inputs);
    return Promise.resolve(ok({}));
  }

  private async getTemplateConfigMetadata(templateId: string) {
    const metadataPath = path.join(getTemplatesFolder(), "metadata", "template-configs.json");
    if (!fs.existsSync(metadataPath)) {
      throw new Error("template-configs.json not found");
    }

    const templates: TemplateConfig[] = await fs.readJson(metadataPath);
    const templateConfig = templates.find((t) => t.id === templateId);
    if (!templateConfig) {
      throw new Error(`Template config for ${templateId} not found`);
    }
    return templateConfig;
  }
}
