// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConditionFunc,
  Inputs,
  IQTreeNode,
  Platform,
  StringValidation,
} from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import {
  apiOperationQuestion,
  apiSpecLocationQuestion,
  SPFxFrameworkQuestion,
  SPFxImportFolderQuestion,
  SPFxPackageSelectQuestion,
  SPFxSolutionQuestion,
  SPFxWebpartNameQuestion,
} from "../../create";
import {
  ApiAuthOptions,
  BotCapabilityOptions,
  MeArchitectureOptions,
  MeCapabilityOptions,
  NotificationBotOptions,
  setTemplateName,
  TabCapabilityOptions,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";

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
          return !inputs[QuestionNames.ApiPluginManifestPath];
        },
        data: apiOperationQuestion(),
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
    condition: { equals: ProjectTypeOptions.botOptionId },
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
    condition: { equals: ProjectTypeOptions.tab().id },
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
    condition: { equals: ProjectTypeOptions.meOptionId },
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
    children: [
      {
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
      },
    ],
  };
}
