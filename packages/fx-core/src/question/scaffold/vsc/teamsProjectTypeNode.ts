// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AppPackageFolderName,
  ConditionFunc,
  DefaultPluginManifestFileName,
  Inputs,
  IQTreeNode,
  OptionItem,
  Platform,
  StringArrayValidation,
  StringValidation,
  UserError,
} from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import {
  apiOperationQuestion,
  apiSpecLocationQuestion,
  apiSpecTypeSelectQuestion,
  searchOpenAPISpecQueryQuestion,
  selectOpenApiSpecQuestion,
  SPFxFrameworkQuestion,
  SPFxImportFolderQuestion,
  SPFxPackageSelectQuestion,
  SPFxSolutionQuestion,
  SPFxWebpartNameQuestion,
} from "../../create";
import { QuestionNames } from "../../questionNames";
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
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import path from "path";
import * as fs from "fs-extra";
import { ODRProvider, ODRServer } from "../../../component/utils/odrProvider";

export function teamsProjectNode(platform: Platform): IQTreeNode {
  return {
    // project-type = Teams Agents and Apps
    condition: { equals: ProjectTypeOptions.teamsOptionId },
    data: {
      name: QuestionNames.TeamsAppType,
      title: getLocalizedString("core.createProjectQuestion.projectType.teamsAgentsAndApps.title"),
      type: "singleSelect",
      staticOptions: [
        TeamsAgentCapabilityOptions.basicChatbot(),
        TeamsAgentCapabilityOptions.customCopilotRag(),
        TeamsAgentCapabilityOptions.collaboratorAgent(),
        TeamsAgentCapabilityOptions.others(),
      ],
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.customCopilot.placeholder"
      ),
      onDidSelection: setTemplateName,
    },
    children: [
      customCopilotRagNode(),
      // aiAgentNode(),
      llmServiceNode({
        enum: [
          TeamsAgentCapabilityOptions.basicChatbot().id,
          TeamsAgentCapabilityOptions.customCopilotRag().id,
        ],
      }),
      azureOpenAINode({ equals: TeamsAgentCapabilityOptions.collaboratorAgent().id }),
      teamsCapabilityNode(platform),
    ],
  };
}

function teamsCapabilityNode(platform: Platform): IQTreeNode {
  return {
    // teams-app-type = Others
    condition: { equals: TeamsAgentCapabilityOptions.others().id },
    data: {
      name: QuestionNames.TeamsCapability,
      title: getLocalizedString("core.createProjectQuestion.teamsCapability.title"),
      type: "singleSelect",
      staticOptions: [
        TabCapabilityOptions.nonSsoTab(),
        MeCapabilityOptions.basicMe(),
        BotCapabilityOptions.basicBot(),
      ],
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.customCopilot.placeholder"
      ),
      onDidSelection: setTemplateName,
    },
    children: [],
  };
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

export function azureOpenAINode(
  condition?: StringValidation | StringArrayValidation | ConditionFunc
): IQTreeNode {
  return {
    condition: condition,
    data: {
      type: "text",
      password: true,
      name: QuestionNames.AzureOpenAIKey,
      title: getLocalizedString("core.createProjectQuestion.llmService.azureOpenAIKey.title"),
      placeholder: getLocalizedString(
        "core.createProjectQuestion.llmService.azureOpenAIKey.placeholder"
      ),
    },
    children: [
      {
        condition: (inputs: Inputs) => {
          return inputs[QuestionNames.AzureOpenAIKey]?.length > 0;
        },
        data: {
          type: "text",
          name: QuestionNames.AzureOpenAIEndpoint,
          title: getLocalizedString(
            "core.createProjectQuestion.llmService.azureOpenAIEndpoint.title"
          ),
          placeholder: getLocalizedString(
            "core.createProjectQuestion.llmService.azureOpenAIEndpoint.placeholder"
          ),
        },
        children: [
          {
            condition: (inputs: Inputs) => {
              return inputs[QuestionNames.AzureOpenAIEndpoint]?.length > 0;
            },
            data: {
              type: "text",
              name: QuestionNames.AzureOpenAIDeploymentName,
              title: getLocalizedString(
                "core.createProjectQuestion.llmService.azureOpenAIDeploymentName.title"
              ),
              placeholder: getLocalizedString(
                "core.createProjectQuestion.llmService.azureOpenAIDeploymentName.placeholder"
              ),
            },
          },
        ],
      },
    ],
  };
}

export function llmServiceNode(
  condition?: StringValidation | StringArrayValidation | ConditionFunc
): IQTreeNode {
  return {
    condition: condition,
    data: {
      type: "singleSelect",
      name: QuestionNames.LLMService,
      title: getLocalizedString("core.createProjectQuestion.llmService.title"),
      placeholder: getLocalizedString("core.createProjectQuestion.llmService.placeholder"),
      staticOptions: [
        {
          id: "llm-service-azure-openai",
          label: getLocalizedString("core.createProjectQuestion.llmServiceAzureOpenAIOption.label"),
          detail: getLocalizedString(
            "core.createProjectQuestion.llmServiceAzureOpenAIOption.detail"
          ),
        },
        {
          id: "llm-service-openai",
          label: getLocalizedString("core.createProjectQuestion.llmServiceOpenAIOption.label"),
          detail: getLocalizedString("core.createProjectQuestion.llmServiceOpenAIOption.detail"),
        },
      ],
      skipSingleOption: true,
      default: "llm-service-azure-openai",
    },
    children: [
      azureOpenAINode({ equals: "llm-service-azure-openai" }),
      {
        condition: { equals: "llm-service-openai" },
        data: {
          type: "text",
          password: true,
          name: QuestionNames.OpenAIKey,
          title: getLocalizedString("core.createProjectQuestion.llmService.openAIKey.title"),
          placeholder: getLocalizedString(
            "core.createProjectQuestion.llmService.openAIKey.placeholder"
          ),
        },
      },
    ],
  };
}

export function apiSpecNode(condition: StringValidation | ConditionFunc): IQTreeNode {
  return {
    condition: condition,
    data: { type: "group", name: QuestionNames.FromExistingApi },
    children: [
      {
        data: apiSpecLocationQuestion(),
      },
      {
        condition: (inputs: Inputs) => {
          return !inputs[QuestionNames.ActionManifestPath];
        },
        data: apiOperationQuestion(),
      },
    ],
  };
}

export function apiSpecWithSearchNode(): IQTreeNode {
  return {
    data: { type: "group", name: QuestionNames.FromExistingApi },
    condition: { equals: "api-spec" },
    children: [inputOrSearchAPISpecNode()],
  };
}

export function inputOrSearchAPISpecNode(): IQTreeNode {
  return {
    data: apiSpecTypeSelectQuestion(),
    condition: (inputs: Inputs) => {
      return featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration);
    },
    children: [
      {
        condition: { equals: "enter-url-or-open-local-file" },
        data: apiSpecLocationQuestion(),
        children: [
          {
            condition: (inputs: Inputs) => {
              return !inputs[QuestionNames.ActionManifestPath];
            },
            data: apiOperationQuestion(true, true),
          },
        ],
      },
      {
        condition: { equals: "search-api" },
        data: searchOpenAPISpecQueryQuestion(),
        children: [
          {
            data: selectOpenApiSpecQuestion(),
          },
          {
            condition: (inputs: Inputs) => {
              return !!inputs[QuestionNames.SelectOpenApiSpec];
            },
            data: apiOperationQuestion(true, true),
          },
        ],
      },
    ],
  };
}

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
      type: "singleSelect",
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
      onDidSelection: (item: string | OptionItem, inputs: Inputs) => {
        try {
          const serverInfo = item as OptionItem;
          const serverData = JSON.parse(serverInfo.data as string);
          inputs[QuestionNames.MCPLocalServerName] = serverInfo.id;
          inputs[QuestionNames.MCPLocalServerIdentifier] = serverData.identifier;
          inputs[QuestionNames.MCPLocalServerCommand] = serverData.command;
          // Store args in the format needed by the template: "arg1", "arg2", "arg3"
          inputs[QuestionNames.MCPLocalServerArgs] = serverData.args
            .map((arg: string) => `"${arg}"`)
            .join(", ");
        } catch {}
      },
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
                  runtime.spec["enable_dynamic_discovery"] === false
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
