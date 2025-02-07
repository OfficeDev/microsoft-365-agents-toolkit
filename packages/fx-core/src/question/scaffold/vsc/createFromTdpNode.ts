// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs, IQTreeNode } from "@microsoft/teamsfx-api";
import { AppDefinition } from "../../../component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import {
  isBot,
  isBotAndBotBasedMessageExtension,
  isBotBasedMessageExtension,
  isPersonalApp,
  needBotCode,
  needTabAndBotCode,
  needTabCode,
} from "../../../component/driver/teamsApp/utils/utils";
import { TemplateNames } from "../../../component/generator/templates/templateNames";
import { QuestionNames } from "../../constants";
import {
  appNameQuestion,
  folderQuestion,
  selectBotIdsQuestion,
  selectTabsContentUrlQuestion,
  selectTabWebsiteUrlQuestion,
} from "../../create";
import { languageNode } from "./createRootNode";

export class CreateFromTdpCapabilityOptions {}

export function getTemplateName(teamsApp: AppDefinition): string | undefined {
  // tab with bot, tab with message extension, tab with bot and message extension
  if (needTabAndBotCode(teamsApp)) {
    return TemplateNames.TabAndDefaultBot;
  }

  // tab only
  if (needTabCode(teamsApp)) {
    return TemplateNames.Tab;
  }

  // bot and message extension
  if (isBotAndBotBasedMessageExtension(teamsApp)) {
    return TemplateNames.BotAndMessageExtension;
  }

  // bot based message extension
  if (isBotBasedMessageExtension(teamsApp)) {
    return TemplateNames.MessageExtension;
  }

  // bot
  if (isBot(teamsApp)) {
    return TemplateNames.DefaultBot;
  }

  return undefined;
}

export function createFromTdpNode(): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) => inputs.teamsAppFromTdp !== undefined,
        data: {
          type: "singleSelect",
          name: QuestionNames.TemplateName,
          title: "Select a template",
          staticOptions: [],
          dynamicOptions: (inputs: Inputs) => {
            const templateName = getTemplateName(inputs.teamsAppFromTdp as AppDefinition);
            if (templateName) return [templateName];
            return [];
          },
          skipSingleOption: true,
        },
      },
      {
        condition: (inputs: Inputs) =>
          inputs.teamsAppFromTdp && isPersonalApp(inputs.teamsAppFromTdp),
        data: { type: "group", name: QuestionNames.RepalceTabUrl },
        children: [
          {
            condition: (inputs: Inputs) =>
              (inputs.teamsAppFromTdp?.staticTabs.filter((o: any) => !!o.websiteUrl) || []).length >
              0,
            data: selectTabWebsiteUrlQuestion(),
          },
          {
            condition: (inputs: Inputs) =>
              (inputs.teamsAppFromTdp?.staticTabs.filter((o: any) => !!o.contentUrl) || []).length >
              0,
            data: selectTabsContentUrlQuestion(),
          },
        ],
      },
      {
        condition: (inputs: Inputs) =>
          inputs.teamsAppFromTdp && needBotCode(inputs.teamsAppFromTdp as AppDefinition),
        data: selectBotIdsQuestion(),
      },
      languageNode(),
      {
        data: folderQuestion(),
      },
      {
        data: appNameQuestion(),
      },
    ],
  };
  return node;
}
