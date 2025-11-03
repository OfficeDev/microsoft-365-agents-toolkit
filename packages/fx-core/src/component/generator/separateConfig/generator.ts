// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConfigFolderName,
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
import os from "os";
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
    const language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    const templateMetadata = getAllTemplatesOnPlatform(inputs.platform).find(
      (t) => t.name === templateName && t.language === language
    );
    const templates = this.readTemplateConfigMetadata();
    const templateConfig = templates.find((t) => t.id === templateMetadata?.id);

    return templateConfig !== undefined;
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
    const templateConfigMetadata = this.getTemplateConfigMetadata(templateMetadata.id);
    await scaffoldConfigComponents(templateConfigMetadata, destinationPath, inputs);
    return Promise.resolve(ok({}));
  }

  private getTemplateConfigMetadata(templateId: string) {
    const templates = this.readTemplateConfigMetadata();
    const templateConfig = templates.find((t) => t.id === templateId);
    if (!templateConfig) {
      throw new Error(`Template config for ${templateId} not found`);
    }
    return templateConfig;
  }

  private readTemplateConfigMetadata(): TemplateConfig[] {
    // template translation files
    let metadataPath = path.join(
      os.homedir(),
      `.${String(ConfigFolderName)}`,
      "metadata",
      "template-configs.json"
    );

    if (!fs.existsSync(metadataPath)) {
      metadataPath = path.join(getTemplatesFolder(), "metadata", "template-configs.json");
    }

    const templates: TemplateConfig[] = fs.readJSONSync(metadataPath);
    return templates;
  }
}
