// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs, IQTreeNode, OptionItem } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
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
import { customEngineAgentProjectTypeNode } from "./customAgentProjectTypeNode";
import { daProjectTypeNode } from "./daProjectTypeNode";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { botProjectTypeNode, meProjectTypeNode, tabProjectTypeNode } from "./teamsProjectTypeNode";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";

export class TdpCapabilityOptions {
  static me(): OptionItem {
    return {
      id: "message-extension",
      label: getLocalizedString("core.MessageExtensionOption.label"),
      description: getLocalizedString("core.MessageExtensionOption.description"),
      detail: getLocalizedString("core.MessageExtensionOption.detail"),
      data: TemplateNames.MessageExtension,
    };
  }
  static botAndMe(): OptionItem {
    return {
      id: "BotAndMessageExtension",
      label: "", // No need to set display name as this option won't be shown in UI
      data: TemplateNames.BotAndMessageExtension,
    };
  }
  static nonSsoTabAndBot(): OptionItem {
    return {
      id: "TabNonSsoAndBot",
      label: "", // No need to set display name as this option won't be shown in UI
      data: TemplateNames.TabAndDefaultBot,
    };
  }
  static nonSsoTab(): OptionItem {
    return {
      id: "tab-non-sso",
      label: `${getLocalizedString("core.TabNonSso.label")}`,
      detail: getLocalizedString("core.TabNonSso.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.Tab,
    };
  }
}

export function getTemplateName(inputs: Inputs): string | undefined {
  if (inputs.teamsAppFromTdp) {
    const teamsApp = inputs.teamsAppFromTdp as AppDefinition;
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
  } else if (
    featureFlagManager.getBooleanValue(FeatureFlags.TdpTemplateCliTest) &&
    inputs.nonInteractive
  ) {
    const capability = inputs[QuestionNames.Capabilities];
    const map = {
      [TdpCapabilityOptions.me().id]: TemplateNames.MessageExtension,
      [TdpCapabilityOptions.nonSsoTab().id]: TemplateNames.Tab,
      [TdpCapabilityOptions.botAndMe().id]: TemplateNames.BotAndMessageExtension,
      [TdpCapabilityOptions.nonSsoTabAndBot().id]: TemplateNames.TabAndDefaultBot,
    };
    const value = map[capability];
    return value;
  }
}

export function isTdpTemplate(inputs: Inputs): boolean {
  const template = getTemplateName(inputs);
  return template !== undefined;
}

export function createFromTdpNode(): IQTreeNode {
  const node: IQTreeNode = {
    condition: isTdpTemplate,
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
        data: {
          name: QuestionNames.ProjectType,
          title: getLocalizedString("core.createProjectQuestion.title"),
          type: "singleSelect",
          staticOptions: [
            ProjectTypeOptions.declarativeAgent(),
            ProjectTypeOptions.customEngineAgent(),
            ProjectTypeOptions.bot(),
            ProjectTypeOptions.tab(),
            ProjectTypeOptions.me(),
          ],
        },
        children: [
          daProjectTypeNode(),
          customEngineAgentProjectTypeNode(),
          botProjectTypeNode(),
          tabProjectTypeNode(),
          meProjectTypeNode(),
        ],
      },
      {
        condition: (inputs: Inputs) =>
          !!inputs.teamsAppFromTdp && isPersonalApp(inputs.teamsAppFromTdp),
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
              !!inputs.teamsAppFromTdp &&
              (inputs.teamsAppFromTdp?.staticTabs.filter((o: any) => !!o.contentUrl) || []).length >
                0,
            data: selectTabsContentUrlQuestion(),
          },
        ],
      },
      {
        condition: (inputs: Inputs) =>
          !!inputs.teamsAppFromTdp && needBotCode(inputs.teamsAppFromTdp as AppDefinition),
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
