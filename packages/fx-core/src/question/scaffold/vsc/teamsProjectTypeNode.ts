// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConditionFunc,
  Inputs,
  IQTreeNode,
  OptionItem,
  Platform,
  StringValidation,
  UserError,
} from "@microsoft/teamsfx-api";
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
  MeArchitectureOptions,
  MeCapabilityOptions,
  NotificationBotOptions,
  setTemplateName,
  TabCapabilityOptions,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { TemplateNames } from "../../../component/generator/templates/templateNames";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";

export function teamsAppProjectNode(platform: Platform): IQTreeNode {
  return {
    // project-type = Teams App
    condition: { equals: ProjectTypeOptions.teamsAppOptionId },
    data: {
      name: QuestionNames.TeamsAppType,
      title: getLocalizedString("core.createProjectQuestion.projectType.teamsApp.title"),
      type: "singleSelect",
      staticOptions: [
        TeamsProjectTypeOptions.bot(platform),
        TeamsProjectTypeOptions.tab(platform),
        TeamsProjectTypeOptions.me(platform),
      ],
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.customCopilot.placeholder"
      ),
    },
    children: [botProjectTypeNode(), tabProjectTypeNode(), meProjectTypeNode()],
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

export function MCPForDAServerUrlNode(): IQTreeNode {
  return {
    condition: { equals: ActionStartOptions.mcp().id },
    data: {
      name: QuestionNames.MCPForDAServerUrl,
      title: getLocalizedString("core.createProjectQuestion.mcpForDa.ServerUrl.title"),
      type: "text",
      placeholder: getLocalizedString("core.createProjectQuestion.mcpForDa.ServerUrl.placeholder"),
    },
    children: [
      {
        data: {
          type: "singleSelect",
          name: QuestionNames.MCPForDATool,
          title: getLocalizedString("core.createProjectQuestion.mcpForDa.Tool.title"),
          staticOptions: [
            {
              id: "pre-fetch-tools",
              label: getLocalizedString("core.createProjectQuestion.mcpForDa.Tool.preFetch.label"),
              detail: getLocalizedString(
                "core.createProjectQuestion.mcpForDa.Tool.preFetch.detail"
              ),
              data: TemplateNames.DeclarativeAgentWithActionFromMCP,
            },
            {
              id: "dynamic-discovery",
              label: getLocalizedString(
                "core.createProjectQuestion.mcpForDa.Tool.dynamicDiscovery.label"
              ),
              detail: getLocalizedString(
                "core.createProjectQuestion.mcpForDa.Tool.dynamicDiscovery.detail"
              ),
              data: TemplateNames.DeclarativeAgentWithActionFromMCP,
            },
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          {
            condition: { equals: "pre-fetch-tools" },
            data: {
              type: "multiSelect",
              name: QuestionNames.MCPForDAPreFetchTools,
              title: getLocalizedString("core.createProjectQuestion.mcpForDa.PreFetchTools.title"),
              staticOptions: [],
              dynamicOptions: async (inputs: Inputs): Promise<OptionItem[]> => {
                const serverUrl = inputs[QuestionNames.MCPForDAServerUrl];
                if (!serverUrl) {
                  return [];
                }
                const transport = new StreamableHTTPClientTransport(new URL(serverUrl));
                const client = new Client({
                  name: "M365 Agent Toolkit",
                  version: "1.0.0",
                });
                try {
                  await client.connect(transport);
                  const mcpTools = await client.listTools();
                  inputs[QuestionNames.MCPForDAPreFetchToolsDetail] = mcpTools.tools;
                  return mcpTools.tools.map((tool: any) => {
                    return {
                      id: tool.name,
                      label: tool.name,
                      detail: tool.description || "",
                    };
                  });
                } catch (error: any) {
                  throw new UserError(
                    "MCPForDAServerUrlError",
                    "Failed to connect to MCP server",
                    `Please check the MCP server URL: ${serverUrl as string}. Error: ${
                      error.message as string
                    }`
                  );
                }
              },
            },
          },
        ],
      },
    ],
  };
}
