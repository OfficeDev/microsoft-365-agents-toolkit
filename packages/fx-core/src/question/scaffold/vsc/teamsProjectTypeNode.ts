// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AppPackageFolderName,
  ConfigFolderName,
  DefaultPluginManifestFileName,
  Inputs,
  IQTreeNode,
  OptionItem,
  Platform,
} from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import os from "os";
import path from "path";
import { getLocalizedString } from "../../../common/localizeUtils";
import { useLocalTemplate } from "../../../component/generator/templateHelper";
import { ODRProvider, ODRServer } from "../../../component/utils/odrProvider";
import { getTemplatesFolder } from "../../../folder";
import {
  SPFxFrameworkQuestion,
  SPFxImportFolderQuestion,
  SPFxPackageSelectQuestion,
  SPFxSolutionQuestion,
  SPFxWebpartNameQuestion,
} from "../../create";
import { QuestionNames } from "../../questionNames";
import { apiSpecNode } from "../commonNodes";
import { constructNode } from "../constructNode";
import {
  ActionStartOptions,
  ApiAuthOptions,
  BotCapabilityOptions,
  CustomCopilotRagOptions,
  MeArchitectureOptions,
  MeCapabilityOptions,
  NotificationBotOptions,
  setTemplateName,
  TabCapabilityOptions,
  TeamsAgentCapabilityOptions,
} from "./CapabilityOptions";

export function getTeamsProjectNode(): IQTreeNode {
  let jsonPath: string;

  const cachedJsonPath = path.join(
    os.homedir(),
    `.${String(ConfigFolderName)}`,
    "ui",
    "teamsNode.json"
  );

  // Check if cached JSON exists, otherwise fallback to bundled templates folder
  if (!useLocalTemplate() && fs.pathExistsSync(cachedJsonPath)) {
    jsonPath = cachedJsonPath;
  } else {
    jsonPath = path.join(getTemplatesFolder(), "ui", "teamsNode.json");
  }

  const content = fs.readFileSync(jsonPath, "utf-8");
  return constructNode(content);
}

export class TeamsProjectTypeOptions {
  static tabOptionId = "tab-type";
  static botOptionId = "bot-type";
  static meOptionId = "me-type";

  static tab(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: TeamsProjectTypeOptions.tabOptionId,
      label: `${platform === Platform.VSCode ? "$(browser) " : ""}${getLocalizedString(
        "core.TabOption.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.tab.detail"),
    };
  }

  static bot(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: TeamsProjectTypeOptions.botOptionId,
      label: `${platform === Platform.VSCode ? "$(hubot) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.bot.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.bot.detail"),
    };
  }

  static me(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: TeamsProjectTypeOptions.meOptionId,
      label: `${platform === Platform.VSCode ? "$(symbol-keyword) " : ""}${getLocalizedString(
        "core.MessageExtensionOption.label"
      )}`,
      detail: getLocalizedString(
        "core.createProjectQuestion.projectType.messageExtension.copilotEnabled.detail"
      ),
    };
  }
}

export function customCopilotRagNode(): IQTreeNode {
  return {
    condition: { equals: TeamsAgentCapabilityOptions.customCopilotRag().id },
    data: {
      type: "singleSelect",
      name: QuestionNames.CustomCopilotRag,
      title: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.label"
      ),
      placeholder: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRag.placeholder"
      ),
      staticOptions: [
        CustomCopilotRagOptions.customize(),
        CustomCopilotRagOptions.azureAISearch(),
        CustomCopilotRagOptions.customApi(),
        // CustomCopilotRagOptions.microsoft365(),
      ],
      default: CustomCopilotRagOptions.customize().id,
      onDidSelection: setTemplateName,
    },
    children: [apiSpecNode({ equals: CustomCopilotRagOptions.customApi().id })],
  };
}

// export function aiAgentNode(): IQTreeNode {
//   return {
//     condition: { equals: CustomCopilotCapabilityOptions.aiAgent().id },
//     data: {
//       type: "singleSelect",
//       name: QuestionNames.CustomCopilotAssistant,
//       title: getLocalizedString(
//         "core.createProjectQuestion.capability.customCopilotAssistant.title"
//       ),
//       placeholder: getLocalizedString(
//         "core.createProjectQuestion.capability.customCopilotAssistant.placeholder"
//       ),
//       staticOptions: [
//         CustomCopilotAssistantOptions.new(),
//         CustomCopilotAssistantOptions.assistantsApi(),
//       ],
//       default: CustomCopilotAssistantOptions.new().id,
//       onDidSelection: setTemplateName,
//     },
//   };
// }

export function notificationBotTriggerNode(platform: Platform = Platform.VSCode): IQTreeNode {
  return {
    condition: { equals: BotCapabilityOptions.notificationBotId },
    data: {
      name: QuestionNames.BotTrigger,
      title: getLocalizedString("plugins.bot.questionHostTypeTrigger.title"),
      type: "singleSelect",
      cliDescription: "Specifies the trigger for `Chat Notification Message` app template.",
      staticOptions: [
        platform === Platform.VS
          ? NotificationBotOptions.appServiceForVS()
          : NotificationBotOptions.appService(),
        NotificationBotOptions.functionsHttpAndTimerTrigger(),
        NotificationBotOptions.functionsHttpTrigger(),
        NotificationBotOptions.functionsTimerTrigger(),
      ],
      default:
        platform === Platform.VS
          ? NotificationBotOptions.appServiceForVS().id
          : NotificationBotOptions.appService().id,
      placeholder: getLocalizedString("plugins.bot.questionHostTypeTrigger.placeholder"),
      onDidSelection: setTemplateName,
    },
  };
}

export function botProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Bot
    condition: { equals: TeamsProjectTypeOptions.botOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.bot.title"),
      type: "singleSelect",
      staticOptions: [
        BotCapabilityOptions.basicBot(),
        BotCapabilityOptions.notificationBot(),
        BotCapabilityOptions.commandBot(),
        BotCapabilityOptions.workflowBot(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [notificationBotTriggerNode()],
  };
}

export function tabProjectTypeNode(platform: Platform = Platform.VSCode): IQTreeNode {
  return {
    // project-type = Tab
    condition: { equals: TeamsProjectTypeOptions.tabOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.tab.title"),
      type: "singleSelect",
      staticOptions: [
        TabCapabilityOptions.nonSsoTab(),
        TabCapabilityOptions.m365SsoLaunchPage(),
        TabCapabilityOptions.dashboardTab(),
        TabCapabilityOptions.SPFxTab(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [
      {
        //SPFx sub-tree
        condition: { equals: TabCapabilityOptions.SPFxTab().id },
        data: SPFxSolutionQuestion(),
        children: [
          {
            data: { type: "group" },
            children: [
              { data: SPFxPackageSelectQuestion() },
              { data: SPFxFrameworkQuestion() },
              { data: SPFxWebpartNameQuestion() },
            ],
            condition: { equals: "new" },
          },
          {
            data: SPFxImportFolderQuestion(),
            condition: { equals: "import" },
          },
        ],
      },
    ],
  };
}

export function meProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Messaging Extension
    condition: { equals: TeamsProjectTypeOptions.meOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.messageExtension.title"),
      type: "singleSelect",
      staticOptions: [
        MeCapabilityOptions.m365SearchMe(),
        MeCapabilityOptions.collectFormMe(),
        MeCapabilityOptions.linkUnfurling(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [m365SearchMeSubNode()],
  };
}

export function m365SearchMeSubNode(): IQTreeNode {
  return {
    // Search ME sub-tree
    condition: { equals: MeCapabilityOptions.m365SearchMe().id },
    data: {
      name: QuestionNames.MeArchitectureType,
      title: getLocalizedString("core.createProjectQuestion.meArchitecture.title"),
      cliDescription: "The authentication type for the API.",
      type: "singleSelect",
      staticOptions: [
        MeArchitectureOptions.newApi(),
        MeArchitectureOptions.openApiSpec(),
        MeArchitectureOptions.botMe(),
      ],
      default: MeArchitectureOptions.newApi().id,
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.copilotExtension.placeholder"
      ),
      forgetLastValue: true,
      skipSingleOption: true,
      onDidSelection: setTemplateName,
    },
    children: [
      {
        condition: { equals: MeArchitectureOptions.newApi().id },
        data: {
          type: "singleSelect",
          name: QuestionNames.ApiAuth,
          title: getLocalizedString("core.createProjectQuestion.apiMessageExtensionAuth.title"),
          placeholder: getLocalizedString(
            "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
          ),
          staticOptions: [
            ApiAuthOptions.none(true),
            ApiAuthOptions.bearerToken(),
            ApiAuthOptions.microsoftEntra(true),
          ],
          default: ApiAuthOptions.none(true).id,
          onDidSelection: setTemplateName,
        },
      },
      apiSpecNode({ equals: MeArchitectureOptions.openApiSpec().id }),
    ],
  };
}

export function MCPServerTypeNode(): IQTreeNode {
  return {
    condition: { equals: ActionStartOptions.mcp().id },
    data: {
      name: QuestionNames.MCPServerType,
      title: getLocalizedString("core.createProjectQuestion.mcpServerType.title"),
      type: "singleSelect",
      staticOptions: [],
      dynamicOptions: async (inputs: Inputs) => {
        const servers = await ODRProvider.listServers();
        inputs["_McpOdrOutput"] = servers;

        const options = [
          {
            id: "remote",
            label: getLocalizedString("core.createProjectQuestion.mcpServerType.remote.label"),
            detail: getLocalizedString("core.createProjectQuestion.mcpServerType.remote.detail"),
          },
        ];

        if (servers.length > 0) {
          options.push({
            id: "local",
            label: getLocalizedString("core.createProjectQuestion.mcpServerType.local.label"),
            detail: getLocalizedString("core.createProjectQuestion.mcpServerType.local.detail"),
          });
        }

        return options;
      },
      default: "remote",
      placeholder: getLocalizedString("core.createProjectQuestion.mcpServerType.placeholder"),
      skipSingleOption: true,
    },
    children: [
      {
        condition: { equals: "remote" },
        data: MCPForDAServerUrlNode().data,
      },
      MCPLocalServerSelectionNode(),
    ],
  };
}

export function MCPLocalServerSelectionNode(): IQTreeNode {
  return {
    condition: { equals: "local" },
    data: {
      name: QuestionNames.MCPLocalServer,
      title: getLocalizedString("core.createProjectQuestion.mcpLocalServer.title"),
      type: "multiSelect",
      returnObject: true,
      staticOptions: [],
      placeholder: getLocalizedString("core.createProjectQuestion.mcpLocalServer.placeholder"),
      dynamicOptions: (inputs: Inputs) => {
        const servers = inputs["_McpOdrOutput"] as ODRServer[];

        return servers.map((server) => ({
          id: server.name,
          label: server.display_name,
          detail: `${server.description} (${server.tools.length} tools available)`,
          data: JSON.stringify({
            identifier: server.identifier,
            command: server.command,
            args: server.args,
          }),
        }));
      },
      validation: { minItems: 1 },
      validationHelp: getLocalizedString(
        "core.createProjectQuestion.mcpLocalServer.validationHelp"
      ),
    },
    children: [],
  };
}

export function MCPForDAServerUrlNode(): IQTreeNode {
  return {
    condition: { equals: ActionStartOptions.mcp().id },
    data: {
      name: QuestionNames.MCPForDAServerUrl,
      title: getLocalizedString("core.createProjectQuestion.mcpForDa.ServerUrl.title"),
      type: "text",
      placeholder: getLocalizedString("core.createProjectQuestion.mcpForDa.ServerUrl.placeholder"),
    },
  };
}

export function updateActionWithMCP(): IQTreeNode {
  return {
    data: {
      type: "singleFile",
      name: QuestionNames.PluginManifestFilePath,
      title: getLocalizedString("core.createProjectQuestion.mcpForDa.File.title"),
      defaultFolder: (inputs: Inputs) => path.normalize(inputs.projectPath as string),
      default: (inputs: Inputs) =>
        path.normalize(
          path.join(
            inputs.projectPath as string,
            AppPackageFolderName,
            DefaultPluginManifestFileName
          )
        ),
    },
    children: [
      {
        data: {
          type: "multiSelect",
          name: QuestionNames.MCPForDAPreFetchTools,
          title: getLocalizedString("core.createProjectQuestion.mcpForDa.PreFetchTools.title"),
          staticOptions: [],
          dynamicOptions: (inputs: Inputs): OptionItem[] => {
            const availableTools: any[] = inputs[QuestionNames.MCPForDAAvailableTools];
            const tools = availableTools.map((tool: any) => {
              return {
                id: tool.name,
                label: tool.name,
                detail: tool.description || "",
              };
            });
            return tools;
          },
          default: async (inputs: Inputs) => {
            const pluginManifestFilePath = inputs[QuestionNames.PluginManifestFilePath];
            if (!pluginManifestFilePath) {
              return [];
            }
            const pluginManifest = await fs.readJSON(pluginManifestFilePath);
            const serverUrl = inputs[QuestionNames.MCPForDAServerUrl];
            const result: string[] = [];
            (pluginManifest.runtimes as any[])
              .filter(
                (runtime: any) =>
                  runtime.type === "RemoteMCPServer" &&
                  runtime.spec.url === serverUrl &&
                  !runtime.spec["enable_dynamic_discovery"]
              )
              .forEach((runtime: any) => {
                result.push(...runtime["run_for_functions"]);
              });
            return result;
          },
        },
      },
      {
        condition: (inputs: Inputs) => {
          return inputs[QuestionNames.MCPForDAAuth] !== "NoneAuth";
        },
        data: {
          type: "singleSelect",
          name: QuestionNames.MCPForDAAuthType,
          title: getLocalizedString("core.createProjectQuestion.mcpForDa.AuthType.title"),
          staticOptions: [
            {
              id: "oauth",
              label: getLocalizedString("core.createProjectQuestion.mcpForDa.Auth.OAuth"),
            },
            {
              id: "entraSSO",
              label: getLocalizedString("core.createProjectQuestion.mcpForDa.Auth.EntraSSO"),
            },
          ],
          default: "oauth",
        },
      },
    ],
  };
}
