// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author zhaofengxu@microsoft.com
 */

import {
  AppPackageFolderName,
  Context,
  err,
  FxError,
  GeneratorResult,
  Inputs,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import { merge } from "lodash";
import path from "path";
import {
  ActionStartOptions,
  ApiAuthOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { developerPortalScaffoldUtils } from "../../developerPortalScaffoldUtils";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { Generator } from "../generator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import fs from "fs-extra";

const enum telemetryProperties {
  templateName = "template-name",
  isDeclarativeCopilot = "is-declarative-copilot",
  isMicrosoftEntra = "is-microsoft-entra",
  needAddPluginFromExisting = "need-add-plugin-from-existing",
}

/**
 * Generator for DA with Copilot connector.
 */
export class CombinedProjectGenerator extends DefaultTemplateGenerator {
  componentName = "combined-project-generator";

  temporaryFolderName = "agent-temp";
  public override activate(context: Context, inputs: Inputs): boolean {
    return [TemplateNames.DeclarativeAgentWithGraphConnector].includes(
      inputs[QuestionNames.TemplateName]
    );
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    const auth = inputs[QuestionNames.ApiAuth];
    const appName = inputs[QuestionNames.AppName];
    const language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    const safeProjectNameFromVS =
      language === "csharp" ? inputs[QuestionNames.SafeProjectName] : undefined;
    const solutionNameFromVS =
      language === "csharp" ? inputs[QuestionNames.SolutionName] : undefined;

    const replaceMap = {
      ...Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        solutionNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true"
      ),
      DeclarativeCopilot: "true",
      CopilotConnector: "true",
      MicrosoftEntra: auth === ApiAuthOptions.microsoftEntra().id ? "true" : "",
    };
    const templateName = inputs[QuestionNames.TemplateName];

    merge(actionContext?.telemetryProps, {
      [telemetryProperties.templateName]: templateName,
      [telemetryProperties.isMicrosoftEntra]:
        auth === ApiAuthOptions.microsoftEntra().id ? "true" : "",
      [telemetryProperties.needAddPluginFromExisting]:
        inputs[QuestionNames.ActionType] === ActionStartOptions.existingPlugin().id.toString(),
    });

    // current template name is declarative agent with Copilot connector, no other template.
    if (templateName === TemplateNames.DeclarativeAgentWithGraphConnector) {
      return Promise.resolve(
        ok([
          {
            templateName: TemplateNames.GraphConnector,
            language: ProgrammingLanguage.TS,
            replaceMap,
          },
          {
            templateName: TemplateNames.DeclarativeAgentBasic,
            language: ProgrammingLanguage.Common,
            replaceMap,
            subFolder: this.temporaryFolderName,
          },
        ])
      );
    }

    return Promise.resolve(
      ok([
        {
          templateName,
          language: language,
          replaceMap,
        },
      ])
    );
  }

  // override this method to do post-step after template download
  async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const srcFolder = path.join(destinationPath, this.temporaryFolderName, AppPackageFolderName);
    const targetFolder = path.join(destinationPath, AppPackageFolderName);
    // copy folder
    fs.copySync(srcFolder, targetFolder, { overwrite: true });
    // delete folder
    fs.removeSync(path.join(destinationPath, this.temporaryFolderName));

    // If coming from TDP portal, apply TDP-specific file updates (e.g. manifest patching)
    if (inputs.teamsAppFromTdp) {
      const res = await developerPortalScaffoldUtils.updateFilesForTdp(
        context,
        inputs.teamsAppFromTdp,
        inputs
      );
      if (res.isErr()) {
        return err(res.error);
      }
    }

    return Promise.resolve(ok({}));
  }
}
