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
  OptionItem,
  Platform,
  Result,
  SystemError,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { merge } from "lodash";
import path from "path";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { convertToAlphanumericOnly } from "../../../common/stringUtils";
import { assembleError } from "../../../error";
import {
  ActionStartOptions,
  ApiAuthOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question";
import { EmbeddedKnowledgeLocalDirectoryName } from "../../driver/teamsApp/constants";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { outputScaffoldingWarningMessage } from "../../utils/common";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { Generator } from "../generator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { setGeneralSensitivityLabel } from "../utils";
import { addExistingPlugin } from "./helper";

const enum telemetryProperties {
  templateName = "template-name",
  isDeclarativeCopilot = "is-declarative-copilot",
  isMicrosoftEntra = "is-microsoft-entra",
  needAddPluginFromExisting = "need-add-plugin-from-existing",
  mcpServerParseFailures = "mcp-server-parse-failures",
  mcpServerParseFailureIds = "mcp-server-parse-failure-ids",
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
      TemplateNames.DeclarativeAgentWithActionFromMCP,
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
    const solutionNameFromVS =
      language === "csharp" ? inputs[QuestionNames.SolutionName] : undefined;

    const MCPServerType = inputs[QuestionNames.MCPServerType];
    const isLocalMCP = MCPServerType === "local";
    const MCPForDAServerUrl = inputs[QuestionNames.MCPForDAServerUrl];

    try {
      const replaceMap = {
        ...Generator.getDefaultVariables(
          inputs[QuestionNames.TemplateName] === TemplateNames.DeclarativeAgentWithTypeSpec
            ? convertToAlphanumericOnly(appName)
            : appName,
          safeProjectNameFromVS,
          solutionNameFromVS,
          inputs.targetFramework,
          inputs.placeProjectFileInSolutionDir === "true"
        ),
        DeclarativeCopilot: "true",
        MicrosoftEntra: auth === ApiAuthOptions.microsoftEntra().id ? "true" : "",
        IsLocalMCP: isLocalMCP ? "true" : "",
        ...(isLocalMCP
          ? this.processMCPLocalServers(inputs, actionContext)
          : MCPForDAServerUrl
          ? {
              MCPForDAServerUrl,
              ServerName: new URL(MCPForDAServerUrl).host
                .replace(/[^a-zA-Z0-9]/g, "")
                .substring(0, 10),
            }
          : {}),
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
    } catch (error) {
      if (error instanceof SystemError) {
        return err(error);
      }
      throw assembleError(error as Error, this.componentName);
    }
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

    // if (
    //   featureFlagManager.getBooleanValue(FeatureFlags.MCPForDA) &&
    //   TemplateNames.DeclarativeAgentWithActionFromMCP === inputs[QuestionNames.TemplateName]
    // ) {
    //   const result = await generateForMCPForDA(destinationPath, inputs);
    //   return result;
    // }

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

  /**
   * Process selected MCP local servers from inputs and format for template
   * Handles both single and multiple server selection for backward compatibility
   */
  private processMCPLocalServers(
    inputs: Inputs,
    context?: ActionContext
  ): {
    MCPLocalServers: Array<{
      name: string;
      identifier: string;
      command: string;
      args: string;
      notLast: boolean;
    }>;
  } {
    const selectedOptions = inputs[QuestionNames.MCPLocalServer] as OptionItem[] | undefined;

    // Handle empty/invalid selection
    if (!selectedOptions || !Array.isArray(selectedOptions) || selectedOptions.length === 0) {
      return {
        MCPLocalServers: [],
      };
    }

    const failures: string[] = [];

    // Map selected options to server configs
    const servers = selectedOptions
      .map((option, index) => {
        try {
          // Validate option structure
          if (!option.data || typeof option.data !== "string") {
            throw new SystemError(
              this.componentName,
              "processMCPLocalServers",
              "Invalid option data structure"
            );
          }

          const serverData = JSON.parse(option.data);

          // Validate parsed data
          if (!serverData.identifier || !serverData.command || !Array.isArray(serverData.args)) {
            throw new SystemError(
              this.componentName,
              "processMCPLocalServers",
              "Invalid server data format"
            );
          }

          return {
            name: option.id,
            identifier: serverData.identifier,
            command: serverData.command,
            args: serverData.args.map((arg: string) => `"${arg}"`).join(", "),
            notLast: index < selectedOptions.length - 1,
          };
        } catch (error) {
          // Track failure for telemetry
          failures.push(option.id);
          return null;
        }
      })
      .filter((server): server is NonNullable<typeof server> => server !== null);

    // If ALL servers failed, throw error
    if (servers.length === 0) {
      throw new SystemError(
        this.componentName,
        "processMCPLocalServers",
        `All ${selectedOptions.length} selected MCP servers failed to parse`
      );
    }

    // Track failures for telemetry
    if (failures.length > 0) {
      merge(context?.telemetryProps, {
        [telemetryProperties.mcpServerParseFailures]: failures.length.toString(),
        [telemetryProperties.mcpServerParseFailureIds]: failures.join(","),
      });
    }

    return {
      MCPLocalServers: servers,
    };
  }
}
