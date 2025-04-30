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
import { addExistingPlugin } from "./helper";
import { getDefaultString } from "../../../common/localizeUtils";
import { EmbeddedKnowledgeLocalDirectoryName } from "../../driver/teamsApp/constants";
import fs from "fs-extra";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { convertToAlphanumericOnly } from "../../../common/stringUtils";
import { setGeneralSensitivityLabel } from "../utils";
import { GraphClient } from "../../../client/graphClient";
import { ListSensitivityLabelScope } from "../../../common/constants";

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
export class DeclarativeAgentGenerator extends DefaultTemplateGenerator {
  componentName = "declarative-agent-from-scratch-generator";
  public override activate(context: Context, inputs: Inputs): boolean {
    return [
      TemplateNames.DeclarativeAgentBasic,
      TemplateNames.DeclarativeAgentWithActionFromScratch,
      TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
      TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
      TemplateNames.DeclarativeAgentWithExistingAction,
      TemplateNames.DeclarativeAgentWithTypeSpec,
    ].includes(inputs[QuestionNames.TemplateName]);
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
        inputs[QuestionNames.TemplateName] === TemplateNames.DeclarativeAgentWithTypeSpec
          ? convertToAlphanumericOnly(appName)
          : appName,
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

  public override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const teamsManifestPath = path.join(
      destinationPath,
      AppPackageFolderName,
      ManifestTemplateFileName
    );
    const declarativeCopilotManifestPathRes = await copilotGptManifestUtils.getManifestPath(
      teamsManifestPath
    );
    if (declarativeCopilotManifestPathRes.isErr()) {
      // only return error in da existing action case
      if (TemplateNames.DeclarativeAgentWithExistingAction === inputs[QuestionNames.TemplateName]) {
        return err(declarativeCopilotManifestPathRes.error);
      }
      return ok({});
    }

    if (featureFlagManager.getBooleanValue(FeatureFlags.SensitivityLabelEnabled)) {
      // best-effort
      await setGeneralSensitivityLabel(context, declarativeCopilotManifestPathRes.value);
    }

    if (
      featureFlagManager.getBooleanValue(FeatureFlags.EmbeddedKnowledgeEnabled) &&
      (inputs.platform === Platform.CLI || inputs.platform === Platform.VSCode)
    ) {
      // ensure EmbeddedKnwoledge folder exists
      const embeddedKnowledgeFolderPath = path.join(
        destinationPath,
        AppPackageFolderName,
        EmbeddedKnowledgeLocalDirectoryName
      );
      await fs.ensureDir(embeddedKnowledgeFolderPath);
    }
    if (TemplateNames.DeclarativeAgentWithExistingAction === inputs[QuestionNames.TemplateName]) {
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
