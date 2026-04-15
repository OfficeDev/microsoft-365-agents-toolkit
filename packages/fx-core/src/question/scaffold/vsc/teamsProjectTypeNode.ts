// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AppPackageFolderName,
  DefaultPluginManifestFileName,
  Inputs,
  IQTreeNode,
  OptionItem,
  Platform,
} from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import path from "path";
import { getLocalizedString } from "../../../common/localizeUtils";
import { fetchMCPTools, readMCPToolsFromFile } from "../../../component/utils/mcpToolFetcher";
import { ODRProvider, ODRServer } from "../../../component/utils/odrProvider";
import {
  SPFxFrameworkQuestion,
  SPFxImportFolderQuestion,
  SPFxPackageSelectQuestion,
  SPFxSolutionQuestion,
  SPFxWebpartNameQuestion,
} from "../../create";
import { QuestionNames } from "../../questionNames";
import { apiSpecNode } from "../commonNodes";
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
import { getRootProjectTypeNode } from "./rootNode";

/**
 * Extract the Teams Agents and Apps sub-tree from the combined wizardNode.json.
 * Used by TDP (Teams Developer Portal) flow.
 */
export function getTeamsProjectNode(): IQTreeNode {
  const root = getRootProjectTypeNode(Platform.VSCode);
  const teamsNode = root.children?.find(
    (c) => (c.condition as any)?.equals === "teams-agent-and-app-type"
  );
  return teamsNode ?? { data: { type: "group" } };
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
      placeholder: getLocalizedString("template.createCapabilityQuestion.placeholder"),
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
      placeholder: getLocalizedString("template.createCapabilityQuestion.placeholder"),
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
      placeholder: getLocalizedString("template.createCapabilityQuestion.placeholder"),
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
        "template.createProjectQuestion.projectType.copilotExtension.placeholder"
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
          title: getLocalizedString("template.createProjectQuestion.apiMessageExtensionAuth.title"),
          placeholder: getLocalizedString(
            "template.createProjectQuestion.apiMessageExtensionAuth.placeholder"
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
        children: [
          MCPToolsFileNode(),
          MCPCliPreFetchToolsNode(),
          {
            condition: (inputs: Inputs) => {
              if (inputs.platform === Platform.VSCode) return false;
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
      },
      MCPLocalServerSelectionNode(),
    ],
  };
}

/**
 * Question node for providing an MCP tools definition file.
 * Shown when:
 * - The MCP server requires auth (auto-detected), or
 * - The user wants to provide tools manually.
 * In CLI non-interactive mode, this is provided via --mcp-tools-file.
 */
export function MCPToolsFileNode(): IQTreeNode {
  return {
    data: {
      name: QuestionNames.MCPToolsFilePath,
      title: getLocalizedString("core.MCPForDA.toolsFilePath.title"),
      type: "text",
      placeholder: getLocalizedString("core.MCPForDA.toolsFilePath.placeholder"),
      additionalValidationOnAccept: {
        validFunc: async (value: string, inputs?: Inputs): Promise<string | undefined> => {
          if (!value) return undefined;
          const filePath = value;
          if (!(await fs.pathExists(filePath))) {
            return getLocalizedString("core.MCPForDA.toolsFileNotFound", filePath);
          }
          try {
            const tools = await readMCPToolsFromFile(filePath);
            if (inputs) {
              inputs[QuestionNames.MCPForDAAvailableTools] = tools;
            }
          } catch (e: any) {
            return e.message;
          }
          return undefined;
        },
      },
    },
    // This node is only shown when:
    // 1. CLI platform (no VS Code MCP gateway), AND
    // 2. Either auth is required, or auto-fetch returned no tools
    condition: (inputs: Inputs) => {
      // In VS Code, tools are fetched by the extension handler, so skip this node
      if (inputs.platform === Platform.VSCode) return false;
      // Show if no available tools were populated by auto-fetch
      const tools = inputs[QuestionNames.MCPForDAAvailableTools];
      return !tools || (Array.isArray(tools) && tools.length === 0);
    },
  };
}

/**
 * Question node for CLI pre-fetch tool selection.
 * Attempts to auto-fetch tools from the MCP server URL, then presents them for selection.
 * For CLI platform only — VS Code has its own tool fetching via vscode.lm.tools.
 */
export function MCPCliPreFetchToolsNode(): IQTreeNode {
  return {
    condition: (inputs: Inputs) => {
      // Only show in CLI when tools are available (either fetched or loaded from file)
      if (inputs.platform === Platform.VSCode) return false;
      const tools = inputs[QuestionNames.MCPForDAAvailableTools];
      return tools && Array.isArray(tools) && tools.length > 0;
    },
    data: {
      type: "multiSelect",
      name: QuestionNames.MCPForDAPreFetchTools,
      title: getLocalizedString("core.createProjectQuestion.mcpForDa.PreFetchTools.title"),
      staticOptions: [],
      dynamicOptions: (inputs: Inputs): OptionItem[] => {
        const availableTools: any[] = inputs[QuestionNames.MCPForDAAvailableTools];
        return availableTools.map((tool: any) => ({
          id: tool.name,
          label: tool.name,
          detail: tool.description || "",
        }));
      },
      default: (inputs: Inputs) => {
        const availableTools: any[] = inputs[QuestionNames.MCPForDAAvailableTools] || [];
        return availableTools.map((tool: any) => tool.name);
      },
    },
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
      additionalValidationOnAccept: {
        validFunc: async (value: string, inputs?: Inputs): Promise<string | undefined> => {
          if (!value || !inputs) return undefined;
          // For CLI: attempt to auto-fetch tools from the server
          if (inputs.platform !== Platform.VSCode) {
            try {
              const result = await fetchMCPTools(value);
              if (result.requiresAuth) {
                inputs["_mcpAuthRequired"] = true;
                inputs[QuestionNames.MCPForDAAvailableTools] = [];
                if (result.authMetadataUrl) {
                  inputs[QuestionNames.MCPForDAAuthMetadataUrl] = result.authMetadataUrl;
                }
                inputs[QuestionNames.MCPForDAAuth] = "OAuthPluginVault";
              } else if (result.tools.length > 0) {
                inputs[QuestionNames.MCPForDAAvailableTools] = result.tools;
                inputs[QuestionNames.MCPForDATool] = "pre-fetch";
                inputs[QuestionNames.MCPForDAAuth] = "NoneAuth";
              } else {
                inputs[QuestionNames.MCPForDAAvailableTools] = [];
                inputs[QuestionNames.MCPForDAAuth] = "NoneAuth";
              }
            } catch {
              inputs[QuestionNames.MCPForDAAvailableTools] = [];
              inputs[QuestionNames.MCPForDAAuth] = "NoneAuth";
            }
          }
          return undefined;
        },
      },
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
                  runtime.type === "RemoteMCPServer" && runtime.spec.url === serverUrl
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
