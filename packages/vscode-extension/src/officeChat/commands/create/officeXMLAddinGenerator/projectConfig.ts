// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

interface IOfficeXMLAddinHostConfig {
  [property: string]: {
    localTemplate: string;
    manifestPath?: string;
    framework: {
      [property: string]: {
        typescript?: string;
        javascript?: string;
      };
    };
  };
}

interface IOfficeXMLAddinProjectConfig {
  [property: string]: IOfficeXMLAddinHostConfig;
}

const CommonProjectConfig = {
  taskpane: {
    framework: {
      default: {
        typescript: "https://aka.ms/ccdevx-fx-taskpane-ts",
        javascript: "https://aka.ms/ccdevx-fx-taskpane-js",
      },
    },
  },
  sso: {
    framework: {
      default: {
        typescript: "https://aka.ms/ccdevx-fx-sso-ts",
        javascript: "https://aka.ms/ccdevx-fx-sso-js",
      },
    },
  },
  react: {
    framework: {
      default: {
        typescript: "https://aka.ms/ccdevx-fx-react-ts",
        javascript: "https://aka.ms/ccdevx-fx-react-js",
      },
    },
  },
  manifest: {
    framework: {
      default: {},
    },
  },
};

export const OfficeXMLAddinProjectConfig: IOfficeXMLAddinProjectConfig = {
  word: {
    "word-taskpane": {
      localTemplate: "word-taskpane",
      ...CommonProjectConfig.taskpane,
    },
    "word-sso": {
      localTemplate: "word-sso",
      ...CommonProjectConfig.sso,
    },
    "word-react": {
      localTemplate: "word-react",
      ...CommonProjectConfig.react,
    },
    "word-manifest": {
      localTemplate: "word-manifest-only",
      ...CommonProjectConfig.manifest,
    },
  },
  excel: {
    "excel-taskpane": {
      localTemplate: "excel-taskpane",
      ...CommonProjectConfig.taskpane,
    },
    "excel-sso": {
      localTemplate: "excel-sso",
      ...CommonProjectConfig.sso,
    },
    "excel-react": {
      localTemplate: "excel-react",
      ...CommonProjectConfig.react,
    },
    "excel-custom-functions-shared": {
      localTemplate: "excel-cf",
      framework: {
        default: {
          typescript: "https://aka.ms/ccdevx-fx-cf-shared-ts",
          javascript: "https://aka.ms/ccdevx-fx-cf-shared-js",
        },
      },
    },
    "excel-custom-functions-js": {
      localTemplate: "excel-cf",
      framework: {
        default: {
          typescript: "https://aka.ms/ccdevx-fx-cf-js-ts",
          javascript: "https://aka.ms/ccdevx-fx-cf-js-js",
        },
      },
    },
    "excel-manifest": {
      localTemplate: "excel-manifest-only",
      ...CommonProjectConfig.manifest,
    },
  },
  powerpoint: {
    "powerpoint-taskpane": {
      localTemplate: "powerpoint-taskpane",
      ...CommonProjectConfig.taskpane,
    },
    "powerpoint-sso": {
      localTemplate: "powerpoint-sso",
      ...CommonProjectConfig.sso,
    },
    "powerpoint-react": {
      localTemplate: "powerpoint-react",
      ...CommonProjectConfig.react,
    },
    "powerpoint-manifest": {
      localTemplate: "powerpoint-manifest-only",
      ...CommonProjectConfig.manifest,
    },
  },
};

export function getOfficeXMLAddinTemplateConfig(addinHost: string): IOfficeXMLAddinHostConfig {
  return OfficeXMLAddinProjectConfig[addinHost];
}
