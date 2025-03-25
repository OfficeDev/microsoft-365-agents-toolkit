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
import { graphAPIClient, listSensitivityLabelScope } from "../../../client/graphAPIClient";
import { getDefaultString } from "../../../common/localizeUtils";
import { EmbeddedKnowledgeLocalDirectoryName } from "../../driver/teamsApp/constants";
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
      await this.setGeneralSensitivityLabel(context, declarativeCopilotManifestPathRes.value);
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

  async setGeneralSensitivityLabel(
    context: Context,
    declarativeAgentManifestPath: string
  ): Promise<void> {
    try {
      const loginStatusRes = await context.tokenProvider?.m365TokenProvider.getStatus({
        scopes: [listSensitivityLabelScope],
      });
      if (!loginStatusRes || loginStatusRes.isErr()) {
        context.logProvider?.info(
          getDefaultString("error.listSensitivityLabel.tokenFailed", loginStatusRes?.error.message)
        );
        return;
      }
      if (loginStatusRes.value.status != signedIn) {
        context.logProvider?.info(getDefaultString("core.listSensitivityLabel.notLogin"));
        return;
      }
      if (loginStatusRes.value.token == undefined) {
        context.logProvider?.info(getDefaultString("error.listSensitivityLabel.tokenUndefined"));
        return;
      }
      const result = await graphAPIClient.getGeneralSentivityLabelId(loginStatusRes.value.token);
      if (result.isErr()) {
        throw result.error;
      }
      const generalLabelId = result.value;

      const declarativeAgentManifestRes = await copilotGptManifestUtils.readCopilotGptManifestFile(
        declarativeAgentManifestPath
      );
      if (declarativeAgentManifestRes.isErr()) {
        context.logProvider?.info(
          getDefaultString(
            "error.readDeclarativeAgentManifest.failed",
            declarativeAgentManifestRes.error
          )
        );
        return;
      }
      const declarativeAgentManifest = declarativeAgentManifestRes.value;
      declarativeAgentManifest.sensitivity_label = generalLabelId;
      const writeRes = await copilotGptManifestUtils.writeCopilotGptManifestFile(
        declarativeAgentManifest,
        declarativeAgentManifestPath
      );
      if (writeRes.isErr()) {
        context.logProvider?.info(
          getDefaultString("error.writeDeclarativeAgentManifest.failed", writeRes.error)
        );
        return;
      }
    } catch (error) {
      context.logProvider?.info(
        getDefaultString("error.setGeneralSensitivityLabel.failed", error.message)
      );
    }
  }
}
