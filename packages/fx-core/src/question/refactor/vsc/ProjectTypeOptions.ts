// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { OptionItem } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";

export class ProjectTypeOptions {
  static tabOptionId = "tab-type";
  static botOptionId = "bot-type";
  static meOptionId = "me-type";
  static outlookAddinOptionId = "outlook-addin-type";
  static officeMetaOSOptionId = "office-meta-os-type";
  static copilotAgentOptionId = "copilot-agent-type";
  static customCopilotOptionId = "custom-copilot-type";
  static startWithGithubCopilotOptionId = "start-with-github-copilot";

  static readonly createGroupName = featureFlagManager.getBooleanValue(
    FeatureFlags.ChatParticipantUIEntries
  )
    ? getLocalizedString("core.createProjectQuestion.projectType.createGroup.title")
    : undefined;

  static tab(): OptionItem {
    return {
      id: ProjectTypeOptions.tabOptionId,
      label: `$(browser) ${getLocalizedString("core.TabOption.label")}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.tab.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static bot(): OptionItem {
    return {
      id: ProjectTypeOptions.botOptionId,
      label: `$(hubot) ${getLocalizedString("core.createProjectQuestion.projectType.bot.label")}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.bot.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static me(): OptionItem {
    return {
      id: ProjectTypeOptions.meOptionId,
      label: `$(symbol-keyword) ${getLocalizedString("core.MessageExtensionOption.label")}`,
      detail: getLocalizedString(
        "core.createProjectQuestion.projectType.messageExtension.copilotEnabled.detail"
      ),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static outlookAddin(): OptionItem {
    return {
      id: ProjectTypeOptions.outlookAddinOptionId,
      label: `$(mail) ${getLocalizedString(
        "core.createProjectQuestion.projectType.outlookAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.outlookAddin.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static officeMetaOS(): OptionItem {
    return {
      id: ProjectTypeOptions.officeMetaOSOptionId,
      label: `$(teamsfx-m365) ${getLocalizedString(
        "core.createProjectQuestion.projectType.officeAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.officeAddin.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static officeAddin(): OptionItem {
    if (featureFlagManager.getBooleanValue(FeatureFlags.OfficeMetaOS)) {
      return this.officeMetaOS();
    } else {
      return this.outlookAddin();
    }
  }

  static Agent(): OptionItem {
    return {
      id: ProjectTypeOptions.copilotAgentOptionId,
      label: `$(teamsfx-agent) ${getLocalizedString(
        "core.createProjectQuestion.projectType.declarativeAgent.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static customCopilot(): OptionItem {
    return {
      id: ProjectTypeOptions.customCopilotOptionId,
      label: `$(teamsfx-custom-copilot) ${getLocalizedString(
        "core.createProjectQuestion.projectType.customCopilot.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.customCopilot.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
  }

  static startWithGithubCopilot(): OptionItem {
    return {
      id: ProjectTypeOptions.startWithGithubCopilotOptionId,
      label: `$(comment-discussion) ${getLocalizedString(
        "core.createProjectQuestion.projectType.copilotHelp.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.copilotHelp.detail"),
      groupName: getLocalizedString("core.createProjectQuestion.projectType.copilotGroup.title"),
    };
  }
}
