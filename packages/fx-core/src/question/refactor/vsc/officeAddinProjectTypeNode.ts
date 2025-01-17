// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import { OfficeAddinCapabilityOptions, setTemplateName } from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";

export function outlookAddinProjectTypeNode(): IQTreeNode {
  return {
    condition: {
      equals: ProjectTypeOptions.outlookAddinOptionId,
    },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.outlookAddin.title"),
      type: "singleSelect",
      staticOptions: [
        OfficeAddinCapabilityOptions.outlookTaskPane(),
        OfficeAddinCapabilityOptions.outlookAddinImport(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      forgetLastValue: true,
      onDidSelection: setTemplateName,
    },
    children: [
      {
        condition: {
          equals: OfficeAddinCapabilityOptions.outlookAddinImport().id,
        },
        data: { type: "group", name: QuestionNames.OfficeAddinImport },
        children: [
          {
            data: {
              type: "folder",
              name: QuestionNames.OfficeAddinFolder,
              title: "Existing add-in project folder",
            },
          },
          {
            data: {
              type: "singleFile",
              name: QuestionNames.OfficeAddinManifest,
              title: "Select import project manifest file",
            },
          },
        ],
      },
    ],
  };
}

export function wxpAddinProjectTypeNode(): IQTreeNode {
  return {
    condition: {
      equals: ProjectTypeOptions.officeMetaOSOptionId,
    },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.officeAddin.title"),
      type: "singleSelect",
      staticOptions: [
        OfficeAddinCapabilityOptions.wxpTaskPane(),
        OfficeAddinCapabilityOptions.officeAddinImport(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      forgetLastValue: true,
      onDidSelection: setTemplateName,
    },
    children: [
      {
        condition: {
          equals: OfficeAddinCapabilityOptions.officeAddinImport().id,
        },
        data: { type: "group", name: QuestionNames.OfficeAddinImport },
        children: [
          {
            data: {
              type: "folder",
              name: QuestionNames.OfficeAddinFolder,
              title: "Existing add-in project folder",
            },
          },
          {
            data: {
              type: "singleFile",
              name: QuestionNames.OfficeAddinManifest,
              title: "Select import project manifest file",
            },
          },
        ],
      },
    ],
  };
}

export function officeAddinProjectTypeNode(): IQTreeNode {
  if (featureFlagManager.getBooleanValue(FeatureFlags.OfficeMetaOS)) {
    return wxpAddinProjectTypeNode();
  } else {
    return outlookAddinProjectTypeNode();
  }
}
