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
} from "@microsoft/teamsfx-api";
import { merge } from "lodash";
import path from "path";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  CapabilityOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { outputScaffoldingWarningMessage } from "../../utils/common";
import { SpecGenerator } from "../apiSpec/generator";
import { declarativeCopilotInstructionFileName } from "../constant";
import { Generator } from "../generator";
import { DefaultTemplateGenerator } from "../templates/templateGenerator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { addExistingPlugin } from "./helper";

const enum telemetryProperties {
  templateName = "template-name",
  isDeclarativeCopilot = "is-declarative-copilot",
  isMicrosoftEntra = "is-microsoft-entra",
  needAddPluginFromExisting = "need-add-plugin-from-existing",
}

/**
 * Generator for copilot extensions including declarative copilot with no plugin,
 * declarative copilot with API plugin from scratch, declarative copilot with existing plugin (to be add later),
 * and API plugin from scratch.
 */
export class CopilotExtensionGenerator extends DefaultTemplateGenerator {
  componentName = "copilot-extension-from-scratch-generator";
  public activate(context: Context, inputs: Inputs): boolean {
    return [
      TemplateNames.ApiPluginFromScratch,
      TemplateNames.ApiPluginFromScratchBearer,
      TemplateNames.ApiPluginFromScratchOAuth,
    ].includes(inputs[QuestionNames.TemplateName]);
    // return (
    //   (inputs[QuestionNames.Capabilities] === CapabilityOptions.declarativeAgent().id &&
    //     inputs[QuestionNames.ApiPluginType] !== ApiPluginStartOptions.apiSpec().id) ||
    //   (inputs[QuestionNames.Capabilities] === CapabilityOptions.apiPlugin().id &&
    //     inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.newApi().id)
    // );
  }

  public getTemplateInfos(
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
    const isDeclarativeCopilot = checkDeclarativeCopilot(inputs);

    const replaceMap = {
      ...Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true"
      ),
      DeclarativeCopilot: isDeclarativeCopilot ? "true" : "",
      MicrosoftEntra: auth === ApiAuthOptions.microsoftEntra().id ? "true" : "",
    };

    const filterFn = (fileName: string) => {
      if (fileName.toLowerCase().includes("declarativeagent.json")) {
        return isDeclarativeCopilot;
      } else if (fileName.includes(declarativeCopilotInstructionFileName)) {
        return isDeclarativeCopilot;
      } else {
        return true;
      }
    };

    // let templateName;
    // const apiPluginFromScratch =
    //   inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.newApi().id;
    // if (apiPluginFromScratch) {
    //   const authTemplateMap = {
    //     [ApiAuthOptions.apiKey().id]: TemplateNames.ApiPluginFromScratchBearer,
    //     [ApiAuthOptions.microsoftEntra().id]: TemplateNames.ApiPluginFromScratchOAuth,
    //     [ApiAuthOptions.oauth().id]: TemplateNames.ApiPluginFromScratchOAuth,
    //   };
    //   templateName = authTemplateMap[auth] || TemplateNames.ApiPluginFromScratch;
    // } else {
    //   templateName = TemplateNames.BasicGpt;
    // }
    const templateName = inputs[QuestionNames.TemplateName];

    merge(actionContext?.telemetryProps, {
      [telemetryProperties.templateName]: templateName,
      [telemetryProperties.isDeclarativeCopilot]: isDeclarativeCopilot.toString(),
      [telemetryProperties.isMicrosoftEntra]:
        auth === ApiAuthOptions.microsoftEntra().id ? "true" : "",
      [telemetryProperties.needAddPluginFromExisting]:
        inputs[QuestionNames.ApiPluginType] ===
        ApiPluginStartOptions.existingPlugin().id.toString(),
    });

    return Promise.resolve(
      ok([
        {
          templateName,
          language: language,
          replaceMap,
          filterFn,
        },
      ])
    );
  }

  public async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const isAddingFromExistingPlugin =
      inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.existingPlugin().id;
    if (isAddingFromExistingPlugin) {
      const teamsManifestPath = path.join(
        destinationPath,
        AppPackageFolderName,
        ManifestTemplateFileName
      );
      const declarativeCopilotManifestPathRes = await copilotGptManifestUtils.getManifestPath(
        teamsManifestPath
      );
      if (declarativeCopilotManifestPathRes.isErr()) {
        return err(declarativeCopilotManifestPathRes.error);
      }
      const addPluginRes = await addExistingPlugin(
        declarativeCopilotManifestPathRes.value,
        inputs[QuestionNames.PluginManifestFilePath],
        inputs[QuestionNames.PluginOpenApiSpecFilePath],
        "action_1",
        context,
        this.componentName
      );

      if (addPluginRes.isErr()) {
        return err(addPluginRes.error);
      } else {
        if (inputs.platform === Platform.CLI || inputs.platform === Platform.VS) {
          const warningMessage = outputScaffoldingWarningMessage(addPluginRes.value.warnings);
          if (warningMessage) {
            context.logProvider.info(warningMessage);
          }
        }
        return ok({ warnings: addPluginRes.value.warnings });
      }
    } else {
      return ok({});
    }
  }
}

function checkDeclarativeCopilot(inputs: Inputs) {
  return inputs[QuestionNames.Capabilities] === CapabilityOptions.declarativeAgent().id;
}

export const specGenerator = new SpecGenerator();
