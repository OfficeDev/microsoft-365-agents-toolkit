// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  AppPackageFolderName,
  Context,
  err,
  FxError,
  GeneratorResult,
  Inputs,
  ManifestTemplateFileName,
  ok,
  Platform,
  Result,
  signedIn,
} from "@microsoft/teamsfx-api";
import { merge } from "lodash";
import path from "path";
import {
  ActionStartOptions,
  ApiAuthOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { outputScaffoldingWarningMessage } from "../../utils/common";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { Generator } from "../generator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { graphAPIClient, listSensitivityLabelScope } from "../../../client/graphAPIClient";
import { getDefaultString } from "../../../common/localizeUtils";
import fs from "fs-extra";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";

const enum telemetryProperties {
  templateName = "template-name",
  isDeclarativeCopilot = "is-declarative-copilot",
  isMicrosoftEntra = "is-microsoft-entra",
  needAddPluginFromExisting = "need-add-plugin-from-existing",
}

/**
 * Generator for copilot extensions including declarative copilot with no plugin,
 * declarative copilot with API plugin from scratch, declarative copilot with existing plugin,
 * and API plugin from scratch.
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

    const replaceMap = {
      ...Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true"
      ),
      DeclarativeCopilot: "true",
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
  protected post(
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
    return Promise.resolve(ok({}));
  }
}
