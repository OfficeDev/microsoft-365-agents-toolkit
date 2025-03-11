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
import {
  graphAPIClient,
  GraphAPIClient,
  listSensitivityLabelScope,
} from "../../../client/graphAPIClient";
import { getSPFxToken } from "../../../common/tools";
import { getDefaultString } from "../../../common/localizeUtils";

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

      // best-effort
      await this.setGeneralSensitivityLabel(context, declarativeCopilotManifestPathRes.value, false);

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
    declarativeAgentManifestPath: string,
    mock = false
  ): Promise<void> {
    let generalLabelId = "";
    if (mock) {
      const res = await graphAPIClient.getGeneralSentivityLabelId("", true);
      if (res.isErr()) {
        return;
      }
      generalLabelId = res.value;
    } else {
      const tokenRes = await context.tokenProvider?.m365TokenProvider.getAccessToken({
        scopes: [listSensitivityLabelScope],
      });
      if (!tokenRes || tokenRes.isErr()) {
        context.logProvider?.info(
          getDefaultString("error.listSensitivityLabel.tokenFailed", tokenRes?.error.message)
        );
        return;
      }
      const result = await graphAPIClient.getGeneralSentivityLabelId(tokenRes.value);
      if (result.isErr()) {
        throw result.error;
      }
      generalLabelId = result.value;
    }
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
  }
}
