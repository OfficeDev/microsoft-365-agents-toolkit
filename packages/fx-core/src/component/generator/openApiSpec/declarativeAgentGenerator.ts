// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author yuqzho@microsoft.com, Ning Tang
 */

import { ProjectType, SpecParserError } from "@microsoft/m365-spec-parser";
import { Context, err, FxError, GeneratorResult, Inputs, Result } from "@microsoft/teamsfx-api";
import { assembleError } from "../../../error";
import { QuestionNames } from "../../../question";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { generateFilesFromApiSpec, getTemplateInfosFromApiSpec } from "./common";
import { convertSpecParserErrorToFxError } from "./helper";

export class DeclarativeAgentWithExistingApiSpecGenerator extends DefaultTemplateGenerator {
  componentName = "da-with-existing-api-generator";

  public override activate(context: Context, inputs: Inputs): boolean {
    return (
      TemplateNames.DeclarativeAgentWithActionFromExistingApiSpec ==
      inputs[QuestionNames.TemplateName]
    );
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    return getTemplateInfosFromApiSpec(context, inputs, ProjectType.Copilot, actionContext);
  }

  public override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    try {
      return await generateFilesFromApiSpec(
        context,
        inputs,
        destinationPath,
        ProjectType.Copilot,
        this.componentName
      );
    } catch (e) {
      let error: FxError;
      if (e instanceof SpecParserError) {
        error = convertSpecParserErrorToFxError(e);
      } else {
        error = assembleError(e);
      }
      return err(error);
    }
  }
}
