// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Context, FxError, Inputs, Result, ok } from "@microsoft/teamsfx-api";
import { CapabilityOptions, ProgrammingLanguage, QuestionNames } from "../../../question/constants";
import { DefaultTemplateGenerator } from "./templateGenerator";
import { TemplateInfo } from "./templateInfo";
import { TemplateNames } from "./templateNames";
import { Generator } from "../generator";

// For the APS.NET server-side rendering tab
export class SsrTabGenerator extends DefaultTemplateGenerator {
  capabilities2TemplateNames = {
    [CapabilityOptions.nonSsoTab().id]: TemplateNames.TabSSR,
    [CapabilityOptions.tab().id]: TemplateNames.SsoTabSSR,
  };
  override activate(context: Context, inputs: Inputs): boolean {
    const capability = inputs.capabilities as string;
    return (
      this.capabilities2TemplateNames[capability] !== undefined &&
      inputs[QuestionNames.ProgrammingLanguage] === ProgrammingLanguage.CSharp &&
      inputs.targetFramework !== "net6.0" &&
      inputs.targetFramework !== "net7.0"
    );
  }
  override getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string
  ): Promise<Result<TemplateInfo[], FxError>> {
    const appName = inputs[QuestionNames.AppName];
    const safeProjectNameFromVS = inputs[QuestionNames.SafeProjectName];
    const isNet8 = inputs.targetFramework === "net8.0";
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
          templateName: this.capabilities2TemplateNames[inputs.capabilities as string],
          language: ProgrammingLanguage.CSharp,
          replaceMap,
        },
      ])
    );
  }
}
