// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";
import { VSCapabilityOptions } from "./scaffold/vs/createRootNode";
import { ProjectTypeOptions } from "./scaffold/vsc/ProjectTypeOptions";
import {
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  DACapabilityOptions,
  MeCapabilityOptions,
  OfficeAddinCapabilityOptions,
  TabCapabilityOptions,
} from "./scaffold/vsc/CapabilityOptions";
import { TdpCapabilityOptions } from "./scaffold/vsc/createFromTdpNode";

export class CapabilityOptions {
  // empty
  static empty(): OptionItem {
    return VSCapabilityOptions.empty();
  }

  // bot
  static basicBot(): OptionItem {
    return BotCapabilityOptions.basicBot();
  }
  static notificationBot(): OptionItem {
    return BotCapabilityOptions.notificationBot();
  }

  static commandBot(): OptionItem {
    return BotCapabilityOptions.commandBot();
  }

  static workflowBot(inputs?: Inputs): OptionItem {
    return BotCapabilityOptions.workflowBot();
  }

  //tab

  static nonSsoTab(): OptionItem {
    return TabCapabilityOptions.nonSsoTab();
  }

  static tab(): OptionItem {
    return VSCapabilityOptions.tab();
  }

  static m365SsoLaunchPage(): OptionItem {
    return TabCapabilityOptions.m365SsoLaunchPage();
  }

  static dashboardTab(): OptionItem {
    return TabCapabilityOptions.dashboardTab();
  }

  static SPFxTab(): OptionItem {
    return TabCapabilityOptions.SPFxTab();
  }

  //message extension
  static linkUnfurling(): OptionItem {
    return MeCapabilityOptions.linkUnfurling();
  }

  static m365SearchMe(): OptionItem {
    return MeCapabilityOptions.m365SearchMe();
  }

  static SearchMe(): OptionItem {
    return VSCapabilityOptions.SearchMeVS();
  }

  static collectFormMe(): OptionItem {
    return MeCapabilityOptions.collectFormMe();
  }
  static me(): OptionItem {
    return TdpCapabilityOptions.me();
  }
  static bots(inputs?: Inputs): OptionItem[] {
    if (inputs?.platform === Platform.VS) {
      return [
        CapabilityOptions.basicBot(),
        CapabilityOptions.aiBot(),
        CapabilityOptions.aiAssistantBot(),
        CapabilityOptions.notificationBot(),
        CapabilityOptions.commandBot(),
        CapabilityOptions.workflowBot(inputs),
      ];
    }
    return [
      CapabilityOptions.basicBot(),
      CapabilityOptions.notificationBot(),
      CapabilityOptions.commandBot(),
      CapabilityOptions.workflowBot(inputs),
    ];
  }

  static tabs(): OptionItem[] {
    return [
      CapabilityOptions.nonSsoTab(),
      CapabilityOptions.m365SsoLaunchPage(),
      CapabilityOptions.dashboardTab(),
      CapabilityOptions.SPFxTab(),
    ];
  }

  static dotnetCaps(inputs?: Inputs): OptionItem[] {
    const capabilities = [
      CapabilityOptions.empty(),
      ...CapabilityOptions.agents(),
      ...CapabilityOptions.customCopilots(inputs),
      ...CapabilityOptions.bots(inputs),
      CapabilityOptions.nonSsoTab(),
      CapabilityOptions.tab(),
      ...CapabilityOptions.collectMECaps(),
    ];
    if (featureFlagManager.getBooleanValue(FeatureFlags.TdpTemplateCliTest)) {
      capabilities.push(CapabilityOptions.me());
    }

    return capabilities;
  }

  /**
   * Collect all capabilities for message extension, including dotnet and nodejs.
   * @returns OptionItem[] capability list
   */
  static collectMECaps(): OptionItem[] {
    return [
      CapabilityOptions.m365SearchMe(),
      CapabilityOptions.collectFormMe(),
      CapabilityOptions.SearchMe(),
      CapabilityOptions.linkUnfurling(),
    ];
  }

  static mes(inputs?: Inputs): OptionItem[] {
    return inputs?.platform === Platform.VS
      ? [
          CapabilityOptions.SearchMe(),
          CapabilityOptions.collectFormMe(),
          CapabilityOptions.linkUnfurling(),
        ]
      : [
          CapabilityOptions.m365SearchMe(),
          CapabilityOptions.collectFormMe(),
          CapabilityOptions.linkUnfurling(),
        ];
  }

  static officeAddinCapabilities(projectType: string): OptionItem[] {
    const isOutlookAddin = projectType === ProjectTypeOptions.outlookAddin().id;
    if (isOutlookAddin) {
      return [
        OfficeAddinCapabilityOptions.outlookTaskPane(),
        OfficeAddinCapabilityOptions.outlookAddinImport(),
      ];
    } else {
      return [
        OfficeAddinCapabilityOptions.wxpTaskPane(),
        OfficeAddinCapabilityOptions.officeAddinImport(),
      ];
    }
  }

  static agents(): OptionItem[] {
    return [CapabilityOptions.declarativeAgent()];
  }

  static customCopilots(inputs?: Inputs): OptionItem[] {
    return inputs?.platform === Platform.VS
      ? [
          CapabilityOptions.customCopilotBasic(),
          CapabilityOptions.customCopilotRag(),
          CapabilityOptions.customCopilotAssistant(),
          // customCopilotWeather is currently only for dotnet
          CapabilityOptions.customCopilotWeather(),
        ]
      : [
          CapabilityOptions.customCopilotBasic(),
          CapabilityOptions.customCopilotRag(),
          CapabilityOptions.customCopilotAssistant(),
        ];
  }

  static tdpIntegrationCapabilities(): OptionItem[] {
    // templates that are used by TDP integration only
    return [
      CapabilityOptions.me(),
      CapabilityOptions.botAndMe(),
      CapabilityOptions.nonSsoTabAndBot(),
    ];
  }

  /**
   * static capability list, which does not depend on any feature flags
   */
  static staticAll(inputs?: Inputs): OptionItem[] {
    const capabilityOptions = [
      CapabilityOptions.empty(),
      ...CapabilityOptions.bots(inputs),
      ...CapabilityOptions.tabs(),
      ...CapabilityOptions.collectMECaps(),
      ...CapabilityOptions.agents(),
      ...CapabilityOptions.customCopilots(),
      ...CapabilityOptions.tdpIntegrationCapabilities(),
    ];
    return capabilityOptions;
  }

  /**
   * dynamic capability list, which depends on feature flags
   */
  static all(inputs?: Inputs): OptionItem[] {
    if (inputs?.platform === Platform.VS) {
      return CapabilityOptions.dotnetCaps(inputs);
    }
    const capabilityOptions = [
      ...CapabilityOptions.bots(inputs),
      ...CapabilityOptions.tabs(),
      ...CapabilityOptions.collectMECaps(),
    ];
    capabilityOptions.push(...CapabilityOptions.agents());
    capabilityOptions.push(...CapabilityOptions.customCopilots());
    if (featureFlagManager.getBooleanValue(FeatureFlags.TdpTemplateCliTest)) {
      // test templates that are used by TDP integration only
      capabilityOptions.push(...CapabilityOptions.tdpIntegrationCapabilities());
    }
    capabilityOptions.push(...CapabilityOptions.officeAddinCapabilities(inputs?.projectType));
    return capabilityOptions;
  }

  static outlookAddinImport(): OptionItem {
    return OfficeAddinCapabilityOptions.outlookAddinImport();
  }

  static officeAddinImport(): OptionItem {
    return OfficeAddinCapabilityOptions.officeAddinImport();
  }

  static officeContentAddin(): OptionItem {
    return OfficeAddinCapabilityOptions.officeContentAddin();
  }

  static nonSsoTabAndBot(): OptionItem {
    return TdpCapabilityOptions.nonSsoTabAndBot();
  }

  static botAndMe(): OptionItem {
    return TdpCapabilityOptions.botAndMe();
  }

  // copilot extension - api plugin
  static apiPlugin(): OptionItem {
    return {
      id: "api-plugin",
      label: getLocalizedString("core.createProjectQuestion.projectType.copilotPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.copilotPlugin.detail"),
    };
  }

  // copilot extension - declarative copilot
  static declarativeAgent(): OptionItem {
    return DACapabilityOptions.declarativeAgent();
  }

  static aiBot(): OptionItem {
    return BotCapabilityOptions.aiBot();
  }

  static aiAssistantBot(): OptionItem {
    return VSCapabilityOptions.aiAssistantBot();
  }

  // custom copilot
  static customCopilotBasic(): OptionItem {
    return CustomCopilotCapabilityOptions.basicChatbot();
  }

  static customCopilotRag(): OptionItem {
    return CustomCopilotCapabilityOptions.customCopilotRag();
  }

  static customCopilotAssistant(): OptionItem {
    return CustomCopilotCapabilityOptions.aiAgent();
  }

  // Currently only for dotnet
  static customCopilotWeather(): OptionItem {
    return VSCapabilityOptions.weatherAgentBot();
  }
}
