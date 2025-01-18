// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode, OptionItem, Platform } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as jsonschema from "jsonschema";
import * as os from "os";
import * as path from "path";
import { ConstantString } from "../../../common/constants";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { createContext } from "../../../common/globalVars";
import { getLocalizedString } from "../../../common/localizeUtils";
import { Constants } from "../../../component/generator/spfx/utils/constants";
import { FileNotFoundError } from "../../../error/common";
import { AppNamePattern, ProgrammingLanguage, QuestionNames } from "../../constants";
import { TemplateNames, Templates } from "../../templates";
import {
  ApiPluginStartOptions,
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  MeCapabilityOptions,
  setTemplateName,
  TabCapabilityOptions,
} from "../vsc/CapabilityOptions";
import { llmServiceNode } from "../vsc/customAgentProjectTypeNode";
import { declarativeAgentProjectTypeNode } from "../vsc/daProjectTypeNode";
import { botTriggerNode } from "../vsc/m365ProjectTypeNode";

export async function getSolutionName(spfxFolder: string): Promise<string | undefined> {
  const yoInfoPath = path.join(spfxFolder, Constants.YO_RC_FILE);
  if (await fs.pathExists(yoInfoPath)) {
    const yoInfo = await fs.readJson(yoInfoPath);
    if (yoInfo["@microsoft/generator-sharepoint"]) {
      return yoInfo["@microsoft/generator-sharepoint"][Constants.YO_RC_SOLUTION_NAME];
    } else {
      return undefined;
    }
  } else {
    throw new FileNotFoundError(Constants.PLUGIN_NAME, yoInfoPath, Constants.IMPORT_HELP_LINK);
  }
}

export function languageNode(): IQTreeNode {
  return {
    condition: (inputs: Inputs) => {
      const templateName = inputs[QuestionNames.TemplateName];
      const languages = Templates.filter((t) => t.name === templateName)
        .map((t) => t.language)
        .filter((lang) => lang !== "none" && lang !== undefined);
      return languages.length > 0;
    },
    data: {
      type: "singleSelect",
      title: getLocalizedString("core.ProgrammingLanguageQuestion.title"),
      name: QuestionNames.ProgrammingLanguage,
      staticOptions: [
        { id: ProgrammingLanguage.JS, label: "JavaScript" },
        { id: ProgrammingLanguage.TS, label: "TypeScript" },
        { id: ProgrammingLanguage.CSharp, label: "C#" },
        { id: ProgrammingLanguage.PY, label: "Python" },
      ],
      dynamicOptions: (inputs: Inputs) => {
        const templateName = inputs[QuestionNames.TemplateName];
        const languages = Templates.filter((t) => t.name === templateName)
          .map((t) => t.language)
          .filter((lang) => lang !== "none" && lang !== undefined);
        return languages;
      },
      skipSingleOption: true,
    },
  };
}

export function folderNode(platform: Platform = Platform.VSCode): IQTreeNode {
  return {
    data: {
      type: "folder",
      name: QuestionNames.Folder,
      title: getLocalizedString("core.question.workspaceFolder.title"),
      cliDescription: "Directory where the project folder will be created in.",
      placeholder: getLocalizedString("core.question.workspaceFolder.placeholder"),
      default:
        platform === Platform.CLI ? "./" : path.join(os.homedir(), ConstantString.RootFolder),
    },
  };
}

export function appNameNode(): IQTreeNode {
  return {
    data: {
      type: "text",
      name: QuestionNames.AppName,
      title: getLocalizedString("core.question.appName.title"),
      default: async (inputs: Inputs) => {
        let defaultName = undefined;
        if (inputs[QuestionNames.SPFxSolution] == "import") {
          defaultName = await getSolutionName(inputs[QuestionNames.SPFxFolder]);
        }
        return defaultName;
      },
      validation: {
        validFunc: async (input: string, previousInputs?: Inputs): Promise<string | undefined> => {
          const schema = {
            pattern: AppNamePattern,
            maxLength: 30,
          };
          if (input.length === 25) {
            // show warning notification because it may exceed the Teams app name max length after appending suffix
            const context = createContext();
            void context.userInteraction.showMessage(
              "warn",
              getLocalizedString("core.QuestionAppName.validation.lengthWarning"),
              false
            );
          }
          const appName = input;
          const validateResult = jsonschema.validate(appName, schema);
          if (validateResult.errors && validateResult.errors.length > 0) {
            if (validateResult.errors[0].name === "pattern") {
              return getLocalizedString("core.QuestionAppName.validation.pattern");
            }
            if (validateResult.errors[0].name === "maxLength") {
              return getLocalizedString("core.QuestionAppName.validation.maxlength");
            }
          }
          if (previousInputs && previousInputs.folder) {
            const folder = previousInputs.folder as string;
            if (folder) {
              const projectPath = path.resolve(folder, appName);
              const exists = await fs.pathExists(projectPath);
              if (exists)
                return getLocalizedString("core.QuestionAppName.validation.pathExist", projectPath);
            }
          }
          return undefined;
        },
      },
      placeholder: getLocalizedString("core.question.appName.placeholder"),
    },
  };
}

export class VSCapabilityOptions {
  // empty
  static empty(): OptionItem {
    return {
      id: "empty",
      label: "Empty",
      data: TemplateNames.Empty,
    };
  }

  static declarativeAgent(): OptionItem {
    return {
      id: "declarative-agent",
      label: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.detail"),
    };
  }
}

/**
 *
 * FxCore API for scaffold: scaffold(questionModel: IQTreeNode, generators: DefaultTemplateGenerator[]): Promise<Result<any, FxError>>
 * Dedicated for VS Code platform
 */

export function scaffoldQuestionForVS(): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        data: {
          name: QuestionNames.Capabilities,
          title: getLocalizedString("core.createCapabilityQuestion.titleNew"),
          type: "singleSelect",
          staticOptions: [
            VSCapabilityOptions.empty(),
            VSCapabilityOptions.declarativeAgent(),
            CustomCopilotCapabilityOptions.customCopilotBasic(),
            CustomCopilotCapabilityOptions.customCopilotRag(),
            CustomCopilotCapabilityOptions.customCopilotAssistant(),
            BotCapabilityOptions.basicBot(),
            BotCapabilityOptions.aiBot(),
            BotCapabilityOptions.aiAssistantBot(),
            BotCapabilityOptions.notificationBot(),
            BotCapabilityOptions.commandBot(),
            BotCapabilityOptions.workflowBot(),
            TabCapabilityOptions.nonSsoTab(),
            TabCapabilityOptions.tab(),
            MeCapabilityOptions.m365SearchMe(),
            MeCapabilityOptions.collectFormMe(),
            MeCapabilityOptions.SearchMeVS(),
            MeCapabilityOptions.linkUnfurling(),
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          declarativeAgentProjectTypeNode(VSCapabilityOptions.declarativeAgent().id),
          llmServiceNode({
            enum: [
              CustomCopilotCapabilityOptions.customCopilotBasic().id,
              CustomCopilotCapabilityOptions.customCopilotRag().id,
              CustomCopilotCapabilityOptions.customCopilotAssistant().id,
            ],
          }),
          botTriggerNode(Platform.VS),
        ],
      },
      languageNode(),
      {
        condition: (inputs: Inputs) => {
          // Only skip this project when need to rediect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec 3. No plugin manifest path
          return !(
            featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
            inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
            inputs[QuestionNames.Capabilities] === VSCapabilityOptions.declarativeAgent().id &&
            !inputs[QuestionNames.ApiPluginManifestPath]
          );
        },
        data: {
          type: "group",
        },
        children: [folderNode(), appNameNode()],
      },
    ],
  };
  return node;
}
