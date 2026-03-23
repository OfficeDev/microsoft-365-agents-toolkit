// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Office Add-in sub-tree for the wizard.
 * Serialized to officeAddinNode.json at build time.
 */
export const officeAddinNode = {
  data: {
    name: "capabilities",
    title: "core.createProjectQuestion.projectType.officeAddin.title",
    type: "singleSelect",
    placeholder: "core.createCapabilityQuestion.placeholder",
    options: [
      {
        id: "wxp-json-taskpane",
        label: "core.newTaskpaneAddin.label",
        detail: "core.newTaskpaneAddin.detail",
        data: "office-addin-wxpo-taskpane",
      },
      {
        id: "wxp-json-cf-shortcut",
        label: "core.newCFShortcut.label",
        detail: "core.newCFShortcut.detail",
        data: "office-addin-excel-cfshortcut",
        featureFlag: "TEAMSFX_CF_SHORTCUT_METAOS",
      },
      {
        id: "office-da-meta-os",
        label: "core.createProjectQuestion.DAMetaOS.label",
        detail: "core.createProjectQuestion.DAMetaOS.detail",
        featureFlag: "TEAMSFX_DA_METAOS",
      },
      {
        id: "office-addin-import",
        label: "core.importOfficeAddin.label",
        detail: "core.importAddin.detail",
        data: "office-addin-config",
      },
    ],
  },
  children: [
    {
      condition: { equals: "office-da-meta-os" },
      data: {
        name: "da-meta-os-capability",
        title: "core.createProjectQuestion.DAMetaOS.capability.title",
        type: "singleSelect",
        placeholder: "core.createCapabilityQuestion.placeholder",
        options: [
          {
            id: "da-meta-os-new-project",
            label: "core.createProjectQuestion.DAMetaOS.capability.newProject.label",
            detail: "core.createProjectQuestion.DAMetaOS.capability.newProject.detail",
            data: "declarative-agent-meta-os-new-project",
          },
          {
            id: "da-meta-os-upgrade-existing-project",
            label: "core.createProjectQuestion.DAMetaOS.capability.upgradeProject.label",
            detail: "core.createProjectQuestion.DAMetaOS.capability.upgradeProject.detail",
            data: "declarative-agent-meta-os-upgrade-project",
          },
        ],
      },
      children: [
        {
          node: "officeAddinFolderNode",
          condition: { equals: "da-meta-os-upgrade-existing-project" },
        },
      ],
    },
    {
      condition: { equals: "office-addin-import" },
      node: "officeAddinImportNode",
    },
  ],
};
