// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  Context,
  FxError,
  GeneratorResult,
  IGenerator,
  Inputs,
  Result,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { merge } from "lodash";
import { TelemetryEvent, TelemetryProperty } from "../../common/telemetry";
import { ProgrammingLanguage, QuestionNames } from "../../question/constants";
import { ProgressMessages, ProgressTitles } from "../messages";
import { ActionContext, ActionExecutionMW } from "../middleware/actionExecutionMW";
import { commonTemplateName, componentName } from "./constant";
import { Generator, templateDefaultOnActionError } from "./generator";
import { GeneratorContext, TemplateActionSeq } from "./generatorAction";
import { getDefaultTemplatesOnPlatform } from "./templates/metadata";
import { TemplateInfo } from "./templates/templateInfo";
import { getTemplateReplaceMap } from "./templates/templateReplaceMap";
import { convertToLangKey, renderTemplateFileData, renderTemplateFileName } from "./utils";

export class DefaultTemplateGenerator implements IGenerator {
  // override this property to send telemetry event with different component name
  componentName = componentName;

  // override this method to determine whether to run this generator
  public activate(context: Context, inputs: Inputs): boolean {
    const templateName = inputs[QuestionNames.TemplateName];
    const platform = inputs.platform;
    const templates = getDefaultTemplatesOnPlatform(platform);
    return templates.some((t) => t.name === templateName);
  }

  // The main entry of the generator. Do not override this method.
  @hooks([
    ActionExecutionMW({
      enableProgressBar: true,
      progressTitle: ProgressTitles.create,
      progressSteps: 1,
      enableTelemetry: true,
      telemetryEventName: TelemetryEvent.GenerateTemplate,
    }),
  ])
  public async run(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const preResult = await this.getTemplateInfos(context, inputs, destinationPath, actionContext);
    if (preResult.isErr()) return err(preResult.error);

    const templateInfos = preResult.value;
    for (const templateInfo of templateInfos) {
      templateInfo.replaceMap = { ...getTemplateReplaceMap(inputs), ...templateInfo.replaceMap };
      await this.scaffolding(context, templateInfo, destinationPath, actionContext);
    }

    const postRes = await this.post(context, inputs, destinationPath, actionContext);
    return postRes;
  }

  // override this method to 1) do pre-step before template download and 2) provide information of templates to be downloaded
  protected getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    // const templateName = getTemplateName(inputs);
    const templateName = inputs[QuestionNames.TemplateName];
    const language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    return Promise.resolve(ok([{ templateName, language }]));
  }

  // override this method to do post-step after template download
  protected post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    return Promise.resolve(ok({}));
  }

  private async scaffolding(
    context: Context,
    templateInfo: TemplateInfo,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<void> {
    const name = templateInfo.templateName;
    const language = convertToLangKey(templateInfo.language) ?? commonTemplateName;
    const replaceMap = templateInfo.replaceMap;
    const filterFn = templateInfo.filterFn ?? (() => true);
    const templateName = `${name}-${language}`;
    merge(actionContext?.telemetryProps, {
      [TelemetryProperty.TemplateName]: templateName,
    });

    const generatorContext: GeneratorContext = {
      name: name,
      language: language,
      destination: destinationPath,
      logProvider: context.logProvider,
      fileNameReplaceFn: (fileName, fileData) =>
        renderTemplateFileName(fileName, fileData, replaceMap)
          .replace(/\\/g, "/")
          .replace(`${name}/`, ""),
      fileDataReplaceFn: (fileName, fileData) =>
        renderTemplateFileData(fileName, fileData, replaceMap),
      filterFn: (fileName) =>
        fileName.replace(/\\/g, "/").startsWith(`${name}/`) && filterFn(fileName),
      onActionError: templateDefaultOnActionError,
    };

    await actionContext?.progressBar?.next(ProgressMessages.generateTemplate);
    context.logProvider.debug(`Downloading app template "${templateName}" to ${destinationPath}`);
    await Generator.generate(generatorContext, TemplateActionSeq);

    merge(actionContext?.telemetryProps, {
      [TelemetryProperty.Fallback]: generatorContext.fallback ? "true" : "false", // Track fallback cases.
    });
  }
}

export const defaultGenerator = new DefaultTemplateGenerator();
