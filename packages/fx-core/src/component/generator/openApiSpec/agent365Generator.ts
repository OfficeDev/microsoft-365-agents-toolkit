// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author KennethBWSong, Ning Tang
 */

import { ProjectType } from "@microsoft/m365-spec-parser";
import { Context, FxError, Inputs, Result, ok } from "@microsoft/teamsfx-api";
import { QuestionNames } from "../../../question";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { Generator } from "../generator";
import { ProgrammingLanguage } from "../../../question/constants";

export class Agent365Generator extends DefaultTemplateGenerator {
  componentName = "agent365-generator";

  public override activate(context: Context, inputs: Inputs): boolean {
    return TemplateNames.Agent365Agents == inputs[QuestionNames.TemplateName];
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    const templateName = inputs[QuestionNames.TemplateName];
    let language = ProgrammingLanguage.TS;
    const llmServiceData = {
      llmService: inputs[QuestionNames.LLMService],
      openAIKey: inputs[QuestionNames.OpenAIKey],
      azureOpenAIKey: inputs[QuestionNames.AzureOpenAIKey],
      azureOpenAIEndpoint: inputs[QuestionNames.AzureOpenAIEndpoint],
      azureOpenAIDeploymentName: inputs[QuestionNames.AzureOpenAIDeploymentName],
    };
    context.templateVariables = Generator.getDefaultVariables(
      inputs[QuestionNames.AppName],
      undefined,
      undefined,
      inputs.targetFramework,
      inputs.placeProjectFileInSolutionDir === "true",
      [],
      llmServiceData
    );
    return ok([
      {
        templateName: templateName,
        language: language,
        replaceMap: {
          ...context.templateVariables,
        },
      },
    ]);
  }
}
