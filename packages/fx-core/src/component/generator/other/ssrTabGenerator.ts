// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Context, FxError, Inputs, Result, ok } from "@microsoft/teamsfx-api";
import { ProgrammingLanguage, QuestionNames } from "../../../question/constants";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { Generator } from "../generator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";

// For the APS.NET server-side rendering tab
export class SsrTabGenerator extends DefaultTemplateGenerator {
  public override activate(context: Context, inputs: Inputs): boolean {
    const templateName = inputs[QuestionNames.TemplateName];
    return [TemplateNames.SsoTabSSR, TemplateNames.TabSSR].includes(templateName);
  }
  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string
  ): Promise<Result<TemplateInfo[], FxError>> {
    const templateName = inputs[QuestionNames.TemplateName];
    const appName = inputs[QuestionNames.AppName];
    const safeProjectNameFromVS = inputs[QuestionNames.SafeProjectName];
    const isNet8 = !inputs.targetFramework || inputs.targetFramework === "net8.0";
    const replaceMap = {
      ...Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true"
      ),
      IsNet8Framework: isNet8 ? "true" : "",
    };

    return Promise.resolve(
      ok([
        {
          templateName: templateName,
          language: ProgrammingLanguage.CSharp,
          replaceMap,
        },
      ])
    );
  }
}
