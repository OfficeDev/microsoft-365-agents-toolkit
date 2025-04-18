// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { OptionItem, Platform } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { ProjectTypeGroup } from "../../constants";

export class ProjectTypeOptions {
  static customEngineAgentOptionId = "custom-engine-agent-type";
  static agentForTeamsOptionId = "agent-for-teams-type";
  static teamsAppOptionId = "teams-app-type";
  static tabOptionId = "tab-type";
  static botOptionId = "bot-type";
  static meOptionId = "me-type";
  static outlookAddinOptionId = "outlook-addin-type";
  static officeMetaOSOptionId = "office-meta-os-type";
  static copilotAgentOptionId = "copilot-agent-type";
  static graphConnectorOptionId = "graph-connector-type";
  static startWithGithubCopilotOptionId = "start-with-github-copilot";

  static groupName(group: ProjectTypeGroup): string | undefined {
    switch (group) {
      case ProjectTypeGroup.AIAgent:
        return getLocalizedString("core.createProjectQuestion.projectType.createGroup.aiAgent");
      case ProjectTypeGroup.M365Apps:
        return getLocalizedString("core.createProjectQuestion.projectType.createGroup.m365Apps");
    }
  }
  static tab(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.tabOptionId,
      label: `${platform === Platform.VSCode ? "$(browser) " : ""}${getLocalizedString(
        "core.TabOption.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.tab.detail"),
      // groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static agentForTeams(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.agentForTeamsOptionId,
      label: `${platform === Platform.VSCode ? "$(hubot) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.agentForTeams.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.agentForTeams.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static teamsApp(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.teamsAppOptionId,
      label: `${platform === Platform.VSCode ? "$(hubot) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.teamsApp.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.teamsApp.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static bot(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.botOptionId,
      label: `${platform === Platform.VSCode ? "$(hubot) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.bot.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.bot.detail"),
      // groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static me(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.meOptionId,
      label: `${platform === Platform.VSCode ? "$(symbol-keyword) " : ""}${getLocalizedString(
        "core.MessageExtensionOption.label"
      )}`,
      detail: getLocalizedString(
        "core.createProjectQuestion.projectType.messageExtension.copilotEnabled.detail"
      ),
      // groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static outlookAddin(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.outlookAddinOptionId,
      label: `${platform === Platform.VSCode ? "$(mail) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.outlookAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.outlookAddin.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static officeMetaOS(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.officeMetaOSOptionId,
      label: `${platform === Platform.VSCode ? "$(teamsfx-m365) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.officeAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.officeAddin.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.M365Apps),
    };
  }

  static officeAddin(platform: Platform = Platform.VSCode): OptionItem {
    if (featureFlagManager.getBooleanValue(FeatureFlags.OfficeMetaOS)) {
      return this.officeMetaOS(platform);
    } else {
      return this.outlookAddin(platform);
    }
  }

  static declarativeAgent(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.copilotAgentOptionId,
      label: `${platform === Platform.VSCode ? "$(teamsfx-agent) " : ""}${getLocalizedString(
        "core.createProjectQuestion.projectType.declarativeAgent.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.AIAgent),
    };
  }

  static customEngineAgent(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.customEngineAgentOptionId,
      label: `${
        platform === Platform.VSCode ? "$(teamsfx-custom-copilot) " : ""
      }${getLocalizedString("core.createProjectQuestion.projectType.customCopilot.label")}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.customCopilot.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.AIAgent),
    };
  }

  static graphConnector(platform: Platform = Platform.VSCode): OptionItem {
    return {
      id: ProjectTypeOptions.graphConnectorOptionId,
      label: `${
        platform === Platform.VSCode ? "$(teamsfx-graph-connector) " : ""
      }${getLocalizedString("core.createProjectQuestion.createGraphConnector.label")}`,
      detail: getLocalizedString("core.createProjectQuestion.createGraphConnector.detail"),
      groupName: ProjectTypeOptions.groupName(ProjectTypeGroup.AIAgent),
    };
  }

  static startWithGithubCopilot(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.HideGitHubCopilotPreviewTag)
      ? undefined
      : getLocalizedString("core.createProjectQuestion.option.description.preview");
    return {
      id: ProjectTypeOptions.startWithGithubCopilotOptionId,
      label: `$(comment-discussion) ${getLocalizedString(
        "core.createProjectQuestion.projectType.copilotHelp.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.copilotHelp.detail"),
      groupName: getLocalizedString("core.createProjectQuestion.projectType.copilotGroup.title"),
      description,
    };
  }
}
