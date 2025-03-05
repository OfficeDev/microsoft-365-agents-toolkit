// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author huajiezhang@microsoft.com
 */

import { Context, err, FxError, GeneratorResult, Inputs, ok, Result } from "@microsoft/teamsfx-api";
import { InputValidationError } from "../../../error";
import { ProgrammingLanguage, QuestionNames } from "../../../question/constants";
import { developerPortalScaffoldUtils } from "../../developerPortalScaffoldUtils";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { Generator } from "../generator";
import { TemplateInfo } from "../templates/templateInfo";

/**
 * TdpGenerator is used to generate code from TDP app definition.
 */
export class TdpGenerator extends DefaultTemplateGenerator {
  componentName = "tdp-generator";

  // activation condition
  public override activate(context: Context, inputs: Inputs): boolean {
    // Reuse some templates which are handled by other generators
    return inputs.teamsAppFromTdp !== undefined;
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    const templateName = inputs[QuestionNames.TemplateName];
    if (!templateName) {
      return err(
        new InputValidationError("teamsAppFromTdp", "Invalid App Definition", "TdpGenerator")
      );
    }
    const appName = inputs[QuestionNames.AppName];
    const safeProjectNameFromVS = inputs[QuestionNames.SafeProjectName];
    const isNet8 = !inputs.targetFramework || inputs.targetFramework === "net8.0"; // used by SSR Tab
    const language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    return Promise.resolve(
      ok([
        {
          templateName: templateName,
          language: language,
          replaceMap: {
            ...Generator.getDefaultVariables(
              appName,
              safeProjectNameFromVS,
              inputs.targetFramework,
              inputs.placeProjectFileInSolutionDir === "true"
            ),
            IsNet8Framework: isNet8 ? "true" : "",
          },
        },
      ])
    );
  }

  public override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const res = await developerPortalScaffoldUtils.updateFilesForTdp(
      context,
      inputs.teamsAppFromTdp,
      inputs
    );
    if (res.isErr()) {
      return err(res.error);
    }
    return ok({});
  }
}
