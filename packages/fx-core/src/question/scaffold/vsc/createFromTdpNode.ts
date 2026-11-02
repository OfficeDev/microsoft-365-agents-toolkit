// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs, IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getAppCapabilities, DevPortalAppResponse } from "../../../component/utils/appInfo";
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
import {
  appNameQuestion,
  folderQuestion,
  selectBotIdsQuestion,
  selectTabsContentUrlQuestion,
  selectTabWebsiteUrlQuestion,
} from "../../create";
import { QuestionNames } from "../../questionNames";
import { languageNode } from "./createRootNode";
import { getTdpProjectTypeNode } from "./rootNode";

export function getTemplateName(inputs: Inputs): string | undefined {
  if (inputs.teamsAppFromTdp) {
    const teamsApp: DevPortalAppResponse = inputs.teamsAppFromTdp;
    // tab with bot
    if (needTabAndBotCode(teamsApp)) {
      return TemplateNames.DefaultBot;
    }

    // tab only
    if (needTabCode(teamsApp)) {
      return TemplateNames.Tab;
    }

    // bot and message extension
    if (isBotAndBotBasedMessageExtension(teamsApp)) {
      return TemplateNames.DefaultBot;
    }

    // bot based message extension, tab with message extension
    if (isBotBasedMessageExtension(teamsApp)) {
      return TemplateNames.DefaultMessageExtension;
    }

    // bot, tab with bot with message extension
    if (isBot(teamsApp)) {
      return TemplateNames.DefaultBot;
    }
  }
}

export function isTdpTemplate(inputs: Inputs): boolean {
  const template = getTemplateName(inputs);
  return template !== undefined;
}

export function createFromTdpNode(platform: Platform = Platform.VSCode): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        // templateName is decided by teamsAppFromTdp itself
        condition: (inputs: Inputs) => getTemplateName(inputs) !== undefined,
        data: {
          type: "singleSelect",
          name: QuestionNames.TemplateName,
          title: "Select a template",
          staticOptions: [],
          dynamicOptions: (inputs: Inputs) => {
            const templateName = getTemplateName(inputs);
            return [templateName!];
          },
          skipSingleOption: true,
        },
      },
      {
        // templateName can not decided by teamsAppFromTdp itself, need user input
        condition: (inputs: Inputs) => getTemplateName(inputs) === undefined,
        ...getTdpProjectTypeNode(platform),
      },
      {
        condition: (inputs: Inputs) =>
          !!inputs.teamsAppFromTdp && isPersonalApp(inputs.teamsAppFromTdp),
        data: { type: "group", name: QuestionNames.RepalceTabUrl },
        children: [
          {
            condition: (inputs: Inputs) =>
              (
                getAppCapabilities(inputs.teamsAppFromTdp).staticTabs?.filter(
                  (o) => !!o.websiteUrl
                ) || []
              ).length > 0 && !needBotCode(inputs.teamsAppFromTdp),
            data: selectTabWebsiteUrlQuestion(),
          },
          {
            condition: (inputs: Inputs) =>
              !!inputs.teamsAppFromTdp &&
              (
                getAppCapabilities(inputs.teamsAppFromTdp).staticTabs?.filter(
                  (o) => !!o.contentUrl
                ) || []
              ).length > 0 &&
              !needBotCode(inputs.teamsAppFromTdp),
            data: selectTabsContentUrlQuestion(),
          },
        ],
      },
      {
        condition: (inputs: Inputs) =>
          !!inputs.teamsAppFromTdp && needBotCode(inputs.teamsAppFromTdp),
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
