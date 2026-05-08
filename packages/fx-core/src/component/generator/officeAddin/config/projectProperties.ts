// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const projectProperties = {
  projectTypes: {
    taskpane: {
      displayname: "template.newTaskpaneAddin.label",
      detail: "template.newTaskpaneAddin.detail",
      manifestPath: "manifest.json",
      templates: {
        typescript: {
          prerelease: "json-preview-yo-office-prerelease",
          repository: "https://github.com/OfficeDev/Office-Addin-TaskPane",
          branch: "json-preview-yo-office",
          archive: "https://aka.ms/teams-toolkit/office-addin-taskpane",

          frameworks: {
            default: {
              repository: "https://github.com/OfficeDev/Office-Addin-TaskPane",
              branch: "json-wxpo-preview",
              prerelease: "json-wxpo-preview",
              archive: "https://aka.ms/teams-toolkit/office-addin-taskpane/ts-default",
            },
            react: {
              repository: "https://github.com/OfficeDev/Office-Addin-TaskPane-React",
              branch: "json-wxpo-preview",
              prerelease: "json-wxpo-preview",
              archive: "https://aka.ms/teams-toolkit/office-addin-taskpane/ts-react",
            },
          },
        },

        javascript: {
          frameworks: {
            default: {
              repository: "https://github.com/OfficeDev/Office-Addin-TaskPane-JS",
              branch: "json-wxpo-preview",
              prerelease: "json-wxpo-preview",
              archive: "https://aka.ms/teams-toolkit/office-addin-taskpane/js-default",
            },
            react: {
              repository: "https://github.com/OfficeDev/Office-Addin-TaskPane-React-JS",
              branch: "json-wxpo-preview",
              prerelease: "json-wxpo-preview",
              archive: "https://aka.ms/teams-toolkit/office-addin-taskpane/js-react",
            },
          },
        },
      },
      supportedHosts: ["Outlook", "Word", "Excel", "PowerPoint"],
    },
  },
  hostTypes: {
    excel: {
      displayname: "Excel",
    },
    onenote: {
      displayname: "OneNote",
    },
    outlook: {
      displayname: "Outlook",
    },
    powerpoint: {
      displayname: "PowerPoint",
    },
    project: {
      displayname: "Project",
    },
    word: {
      displayname: "Word",
    },
  },
};
