// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AppPackageFolderName,
  BuildFolderName,
  ConfirmQuestion,
  DynamicPlatforms,
  IQTreeNode,
  Inputs,
  ManifestUtil,
  MultiSelectQuestion,
  Platform,
  SingleFileQuestion,
  SingleSelectQuestion,
  TextInputQuestion,
  FolderQuestion,
  CLIPlatforms,
  PluginManifestSchema,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import { AppStudioScopes, ConstantString } from "../common/constants";
import { FeatureFlags, featureFlagManager } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";
import { Constants } from "../component/driver/add/utility/constants";
import { envUtil } from "../component/utils/envUtil";
import { CollaborationConstants, CollaborationUtil } from "../core/collaborator";
import { environmentNameManager } from "../core/environmentName";
import { TOOLS } from "../common/globalVars";
import {
  AddAuthActionAuthTypeOptions,
  ApiPluginStartOptions,
  HubOptions,
  QuestionNames,
  TeamsAppValidationOptions,
} from "./constants";
import {
  SPFxFrameworkQuestion,
  SPFxImportFolderQuestion,
  SPFxWebpartNameQuestion,
  apiAuthQuestion,
  apiOperationQuestion,
  apiPluginStartQuestion,
  apiSpecLocationQuestion,
  pluginApiSpecQuestion,
  pluginManifestQuestion,
  addKnowledgeStartQuestion,
} from "./create";
import { UninstallInputs } from "./inputs";
import * as os from "os";

export function listCollaboratorQuestionNode(): IQTreeNode {
  const selectTeamsAppNode = selectTeamsAppManifestQuestionNode();
  selectTeamsAppNode.condition = { contains: CollaborationConstants.TeamsAppQuestionId };
  selectTeamsAppNode.children!.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  const selectAadAppNode = selectAadAppManifestQuestionNode();
  selectAadAppNode.condition = { contains: CollaborationConstants.AadAppQuestionId };
  selectAadAppNode.children!.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  return {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) => DynamicPlatforms.includes(inputs.platform),
        data: selectAppTypeQuestion(),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
        children: [selectTeamsAppNode, selectAadAppNode],
      },
    ],
  };
}

export function grantPermissionQuestionNode(): IQTreeNode {
  const selectTeamsAppNode = selectTeamsAppManifestQuestionNode();
  selectTeamsAppNode.condition = { contains: CollaborationConstants.TeamsAppQuestionId };
  selectTeamsAppNode.children!.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  const selectAadAppNode = selectAadAppManifestQuestionNode();
  selectAadAppNode.condition = { contains: CollaborationConstants.AadAppQuestionId };
  selectAadAppNode.children!.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  return {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) => DynamicPlatforms.includes(inputs.platform),
        data: selectAppTypeQuestion(),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
        children: [
          selectTeamsAppNode,
          selectAadAppNode,
          {
            data: inputUserEmailQuestion(),
          },
        ],
      },
    ],
  };
}

export function convertAadToNewSchemaQuestionNode(): IQTreeNode {
  return {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) =>
          DynamicPlatforms.includes(inputs.platform) &&
          !inputs[QuestionNames.AadAppManifestFilePath],
        data: selectAadManifestQuestion(),
        children: [
          {
            condition: (inputs: Inputs) =>
              inputs.platform === Platform.VSCode && // confirm question only works for VSC
              inputs.projectPath !== undefined &&
              path.resolve(inputs[QuestionNames.AadAppManifestFilePath]) !==
                path.join(inputs.projectPath, "aad.manifest.json"),
            data: confirmManifestQuestion(false, false),
            cliOptionDisabled: "self",
            inputsDisabled: "self",
          },
        ],
      },
    ],
  };
}

export function deployAadManifestQuestionNode(): IQTreeNode {
  return {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) => DynamicPlatforms.includes(inputs.platform),
        data: selectAadManifestQuestion(),
        children: [
          {
            condition: (inputs: Inputs) =>
              inputs.platform === Platform.VSCode && // confirm question only works for VSC
              inputs.projectPath !== undefined &&
              path.resolve(inputs[QuestionNames.AadAppManifestFilePath]) !==
                path.join(inputs.projectPath, "aad.manifest.json"),
            data: confirmManifestQuestion(false, false),
            cliOptionDisabled: "self",
            inputsDisabled: "self",
          },
          {
            condition: isAadMainifestContainsPlaceholder,
            data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
          },
        ],
      },
    ],
  };
}

export function validateTeamsAppQuestionNode(): IQTreeNode {
  return {
    data: selectTeamsAppValidationMethodQuestion(),
    cliOptionDisabled: "self",
    inputsDisabled: "self",
    children: [
      {
        condition: { equals: TeamsAppValidationOptions.schema().id },
        data: selectTeamsAppManifestQuestion(),
      },
      {
        condition: { equals: TeamsAppValidationOptions.package().id },
        data: selectTeamsAppPackageQuestion(),
      },
      {
        condition: { equals: TeamsAppValidationOptions.testCases().id },
        data: selectTeamsAppPackageQuestion(),
      },
    ],
  };
}

export function selectTeamsAppManifestQuestionNode(): IQTreeNode {
  return {
    data: selectTeamsAppManifestQuestion(),
    children: [
      {
        condition: (inputs: Inputs) => confirmCondition(inputs, false),
        data: confirmManifestQuestion(true, false),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
      },
    ],
  };
}

export function selectAadAppManifestQuestionNode(): IQTreeNode {
  return {
    data: selectAadManifestQuestion(),
    children: [
      {
        condition: (inputs: Inputs) =>
          inputs.platform === Platform.VSCode && // confirm question only works for VSC
          inputs.projectPath &&
          inputs[QuestionNames.AadAppManifestFilePath] &&
          path.resolve(inputs[QuestionNames.AadAppManifestFilePath]) !==
            path.join(inputs.projectPath, "aad.manifest.json"),
        data: confirmManifestQuestion(false, false),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
      },
    ],
  };
}

function confirmCondition(inputs: Inputs, isLocal: boolean): boolean {
  return (
    inputs.platform === Platform.VSCode && // confirm question only works for VSC
    inputs.projectPath &&
    inputs[
      isLocal ? QuestionNames.LocalTeamsAppManifestFilePath : QuestionNames.TeamsAppManifestFilePath
    ] &&
    path.resolve(
      inputs[
        isLocal
          ? QuestionNames.LocalTeamsAppManifestFilePath
          : QuestionNames.TeamsAppManifestFilePath
      ]
    ) !==
      path.join(
        inputs.projectPath,
        AppPackageFolderName,
        isLocal ? "manifest.local.json" : "manifest.json"
      )
  );
}

async function spfxFrameworkExist(inputs: Inputs): Promise<boolean> {
  if (inputs.platform === Platform.CLI_HELP) {
    return false;
  }

  const yorcPath = path.join(inputs[QuestionNames.SPFxFolder], Constants.YO_RC_FILE);
  if (!(await fs.pathExists(yorcPath))) {
    return false;
  }

  const yorcJson = (await fs.readJson(yorcPath)) as Record<string, any>;
  if (!yorcJson["@microsoft/generator-sharepoint"]) {
    return false;
  }

  return yorcJson["@microsoft/generator-sharepoint"]["template"];
}

export function addWebPartQuestionNode(): IQTreeNode {
  return {
    data: SPFxImportFolderQuestion(true),
    children: [
      {
        data: SPFxWebpartNameQuestion(),
        children: [
          {
            data: SPFxFrameworkQuestion(),
            condition: async (inputs: Inputs) => {
              return !(await spfxFrameworkExist(inputs));
            },
          },
          {
            data: selectTeamsAppManifestQuestion(),
            children: [
              {
                condition: (inputs: Inputs) => confirmCondition(inputs, false),
                data: confirmManifestQuestion(true, false),
                cliOptionDisabled: "self",
                inputsDisabled: "self",
              },
              {
                data: selectLocalTeamsAppManifestQuestion(),
                children: [
                  {
                    condition: (inputs: Inputs) => confirmCondition(inputs, true),
                    data: confirmManifestQuestion(true, true),
                    cliOptionDisabled: "self",
                    inputsDisabled: "self",
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

export function selectTeamsAppManifestQuestion(): SingleFileQuestion {
  return {
    name: QuestionNames.TeamsAppManifestFilePath,
    cliName: "teams-manifest-file",
    cliShortName: "t",
    cliDescription:
      "Specify the path for Teams app manifest template. It can be either absolute path or relative path to the project root folder, with default at './appPackage/manifest.json'",
    title: getLocalizedString("core.selectTeamsAppManifestQuestion.title"),
    type: "singleFile",
    default: (inputs: Inputs): string | undefined => {
      if (inputs.platform === Platform.CLI_HELP) {
        return "./appPackage/manifest.json";
      } else {
        if (!inputs.projectPath) return undefined;
        const manifestPath = path.join(inputs.projectPath, AppPackageFolderName, "manifest.json");
        if (fs.pathExistsSync(manifestPath)) {
          return manifestPath;
        } else {
          return undefined;
        }
      }
    },
  };
}

export function selectLocalTeamsAppManifestQuestion(): SingleFileQuestion {
  return {
    name: QuestionNames.LocalTeamsAppManifestFilePath,
    cliName: "local-teams-manifest-file",
    cliShortName: "l",
    cliDescription:
      "Specifies the Microsoft Teams app manifest template file path for local environment, it can be either absolute path or relative path to project root folder.",
    title: getLocalizedString("core.selectLocalTeamsAppManifestQuestion.title"),
    type: "singleFile",
    default: (inputs: Inputs): string | undefined => {
      if (inputs.platform === Platform.CLI_HELP) {
        return "./appPackage/manifest.local.json";
      } else {
        if (!inputs.projectPath) return undefined;
        const manifestPath = path.join(
          inputs.projectPath,
          AppPackageFolderName,
          "manifest.local.json"
        );
        if (fs.pathExistsSync(manifestPath)) {
          return manifestPath;
        } else {
          return undefined;
        }
      }
    },
  };
}

function confirmManifestQuestion(isTeamsApp = true, isLocal = false): SingleSelectQuestion {
  const map: Record<string, string> = {
    true_true: QuestionNames.ConfirmLocalManifest,
    true_false: QuestionNames.ConfirmManifest,
    false_true: QuestionNames.ConfirmAadManifest,
    false_false: QuestionNames.ConfirmAadManifest,
  };
  const name = map[`${isTeamsApp.toString()}_${isLocal.toString()}`];
  return {
    name: name,
    title: isTeamsApp
      ? getLocalizedString(
          isLocal
            ? "core.selectLocalTeamsAppManifestQuestion.title"
            : "core.selectTeamsAppManifestQuestion.title"
        )
      : getLocalizedString("core.selectAadAppManifestQuestion.title"),
    type: "singleSelect",
    staticOptions: [],
    skipSingleOption: false,
    placeholder: getLocalizedString("core.confirmManifestQuestion.placeholder"),
    dynamicOptions: (inputs: Inputs) => {
      return [
        {
          id: "manifest",
          label: `$(file) ${path.basename(
            isTeamsApp
              ? inputs[
                  isLocal
                    ? QuestionNames.LocalTeamsAppManifestFilePath
                    : QuestionNames.TeamsAppManifestFilePath
                ]
              : inputs[QuestionNames.AadAppManifestFilePath]
          )}`,
          description: path.dirname(
            isTeamsApp
              ? inputs[
                  isLocal
                    ? QuestionNames.LocalTeamsAppManifestFilePath
                    : QuestionNames.TeamsAppManifestFilePath
                ]
              : inputs[QuestionNames.AadAppManifestFilePath]
          ),
        },
      ];
    },
  };
}

function selectTeamsAppValidationMethodQuestion(): SingleSelectQuestion {
  const options = [TeamsAppValidationOptions.schema(), TeamsAppValidationOptions.package()];

  if (featureFlagManager.getBooleanValue(FeatureFlags.AsyncAppValidation)) {
    options.push(TeamsAppValidationOptions.testCases());
  }

  return {
    name: QuestionNames.ValidateMethod,
    title: getLocalizedString("core.selectValidateMethodQuestion.validate.selectTitle"),
    staticOptions: options,
    type: "singleSelect",
  };
}

export function copilotPluginAddAPIQuestionNode(): IQTreeNode {
  return {
    data: apiSpecLocationQuestion(false),
    children: [
      {
        data: apiOperationQuestion(false),
      },
    ],
  };
}

function selectTeamsAppPackageQuestion(): SingleFileQuestion {
  return {
    name: QuestionNames.TeamsAppPackageFilePath,
    title: getLocalizedString("core.selectTeamsAppPackageQuestion.title"),
    cliDescription:
      "Specifies the zipped Microsoft Teams app package path, it's a relative path to project root folder, defaults to '${folder}/appPackage/build/appPackage.${env}.zip'",
    cliName: "app-package-file",
    cliShortName: "p",
    type: "singleFile",
    default: (inputs: Inputs): string | undefined => {
      if (!inputs.projectPath) return undefined;
      const appPackagePath: string = path.join(
        inputs.projectPath,
        AppPackageFolderName,
        BuildFolderName,
        "appPackage.dev.zip"
      );
      if (fs.pathExistsSync(appPackagePath)) {
        return appPackagePath;
      } else {
        return undefined;
      }
    },
  };
}

export function selectTeamsAppPackageQuestionNode(): IQTreeNode {
  return {
    data: selectTeamsAppPackageQuestion(),
  };
}

function selectM365HostQuestion(): SingleSelectQuestion {
  return {
    name: QuestionNames.M365Host,
    cliShortName: "m",
    cliDescription: "Preview the application in Teams, Outlook or the Microsoft 365 app.",
    title: getLocalizedString("core.M365HostQuestion.title"),
    default: HubOptions.teams().id,
    type: "singleSelect",
    staticOptions: HubOptions.all(),
    placeholder: getLocalizedString("core.M365HostQuestion.placeholder"),
  };
}

export function previewWithTeamsAppManifestQuestionNode(): IQTreeNode {
  return {
    data: { type: "group" },
    children: [
      {
        data: selectM365HostQuestion(),
      },
      selectTeamsAppManifestQuestionNode(),
    ],
  };
}

export function selectTargetEnvQuestion(
  questionName = QuestionNames.TargetEnvName,
  remoteOnly = true,
  throwErrorIfNoEnv = false,
  defaultValueIfNoEnv = environmentNameManager.getDefaultEnvName()
): SingleSelectQuestion {
  return {
    type: "singleSelect",
    name: questionName,
    title: getLocalizedString("core.QuestionSelectTargetEnvironment.title"),
    cliName: "env",
    cliDescription: "Specifies the environment name for the project.",
    staticOptions: [],
    dynamicOptions: async (inputs: Inputs) => {
      if (!inputs.projectPath) return [];
      const res = await envUtil.listEnv(inputs.projectPath, remoteOnly);
      if (res.isErr()) {
        if (throwErrorIfNoEnv) throw res.error;
        return [defaultValueIfNoEnv];
      }
      // "testtool" env is a pure local env and doesn't have manifest
      return res.value.filter((env) => env !== environmentNameManager.getTestToolEnvName());
    },
    skipSingleOption: true,
    forgetLastValue: true,
  };
}

async function getDefaultUserEmail() {
  if (!TOOLS?.tokenProvider.m365TokenProvider) return undefined;
  const jsonObjectRes = await TOOLS.tokenProvider.m365TokenProvider.getJsonObject({
    scopes: AppStudioScopes,
  });
  if (jsonObjectRes.isErr()) {
    throw jsonObjectRes.error;
  }
  const jsonObject = jsonObjectRes.value;
  const currentUserEmail = (jsonObject as any).upn as string;
  let defaultUserEmail = "";
  if (currentUserEmail && currentUserEmail.indexOf("@") > 0) {
    defaultUserEmail = "[UserName]@" + currentUserEmail.split("@")[1];
  }
  return defaultUserEmail;
}

export function inputUserEmailQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.UserEmail,
    type: "text",
    title: getLocalizedString("core.getUserEmailQuestion.title"),
    cliDescription: "Email address of the collaborator.",
    default: getDefaultUserEmail,
    validation: {
      validFunc: async (input: string, previousInputs?: Inputs) => {
        if (!input || input.trim() === "") {
          return getLocalizedString("core.getUserEmailQuestion.validation1");
        }

        input = input.trim();
        const defaultUserEmail = await getDefaultUserEmail();
        if (input === defaultUserEmail) {
          return getLocalizedString("core.getUserEmailQuestion.validation2");
        }

        const re = /\S+@\S+\.\S+/;
        if (!re.test(input)) {
          return getLocalizedString("core.getUserEmailQuestion.validation3");
        }
        return undefined;
      },
    },
  };
}

export async function isAadMainifestContainsPlaceholder(inputs: Inputs): Promise<boolean> {
  const aadManifestPath = inputs?.[QuestionNames.AadAppManifestFilePath];
  const placeholderRegex = /\$\{\{ *[a-zA-Z0-9_.-]* *\}\}/g;
  const regexObj = new RegExp(placeholderRegex);
  try {
    if (!aadManifestPath || !(await fs.pathExists(aadManifestPath))) {
      return false;
    }
    const manifest = await fs.readFile(aadManifestPath, ConstantString.UTF8Encoding);
    if (regexObj.test(manifest)) {
      return true;
    }
  } catch (e) {
    return false;
  }
  return false;
}

export function selectAadManifestQuestion(): SingleFileQuestion {
  return {
    name: QuestionNames.AadAppManifestFilePath,
    cliName: "entra-app-manifest-file",
    cliShortName: "a",
    cliDescription:
      "Specifies the Microsoft Entra app manifest file path, can be either absolute path or relative path to project root folder.",
    title: getLocalizedString("core.selectAadAppManifestQuestion.title"),
    type: "singleFile",
    default: (inputs: Inputs): string | undefined => {
      if (inputs.platform === Platform.CLI_HELP) {
        return "./aad.manifest.json";
      } else {
        if (!inputs.projectPath) return undefined;
        const manifestPath: string = path.join(inputs.projectPath, "aad.manifest.json");
        if (fs.pathExistsSync(manifestPath)) {
          return manifestPath;
        } else {
          return undefined;
        }
      }
    },
  };
}

function selectAppTypeQuestion(): MultiSelectQuestion {
  return {
    name: QuestionNames.collaborationAppType,
    title: getLocalizedString("core.selectCollaborationAppTypeQuestion.title"),
    type: "multiSelect",
    staticOptions: [
      {
        id: CollaborationConstants.AadAppQuestionId,
        label: getLocalizedString("core.aadAppQuestion.label"),
        description: getLocalizedString("core.aadAppQuestion.description"),
      },
      {
        id: CollaborationConstants.TeamsAppQuestionId,
        label: getLocalizedString("core.teamsAppQuestion.label"),
        description: getLocalizedString("core.teamsAppQuestion.description"),
      },
    ],
    validation: { minItems: 1 },
    validationHelp: "Please select at least one app type.",
  };
}

export async function envQuestionCondition(inputs: Inputs): Promise<boolean> {
  const appType = inputs[CollaborationConstants.AppType] as string[];
  const requireAad = appType?.includes(CollaborationConstants.AadAppQuestionId);
  const requireTeams = appType?.includes(CollaborationConstants.TeamsAppQuestionId);
  const aadManifestPath = inputs[QuestionNames.AadAppManifestFilePath];
  const teamsManifestPath = inputs[QuestionNames.TeamsAppManifestFilePath];

  // When both is selected, only show the question once at the end
  if ((requireAad && !aadManifestPath) || (requireTeams && !teamsManifestPath)) {
    return false;
  }

  // Only show env question when manifest id is referencing value from .env file
  let requireEnv = false;
  if (requireTeams && teamsManifestPath) {
    const teamsAppIdRes = await CollaborationUtil.loadManifestId(teamsManifestPath);
    if (teamsAppIdRes.isOk()) {
      requireEnv = CollaborationUtil.requireEnvQuestion(teamsAppIdRes.value);
      if (requireEnv) {
        return true;
      }
    } else {
      return false;
    }
  }

  if (requireAad && aadManifestPath) {
    const aadAppIdRes = await CollaborationUtil.loadManifestId(aadManifestPath);
    if (aadAppIdRes.isOk()) {
      requireEnv = CollaborationUtil.requireEnvQuestion(aadAppIdRes.value);
      if (requireEnv) {
        return true;
      }
    } else {
      return false;
    }
  }

  return false;
}
export async function newEnvNameValidation(
  input: string,
  inputs?: Inputs
): Promise<string | undefined> {
  const targetEnvName = input;
  const match = targetEnvName.match(environmentNameManager.envNameRegex);
  if (!match) {
    return getLocalizedString("core.getQuestionNewTargetEnvironmentName.validation1");
  }

  if (!environmentNameManager.isRemoteEnvironment(targetEnvName)) {
    return getLocalizedString(
      "core.getQuestionNewTargetEnvironmentName.validation3",
      targetEnvName
    );
  }
  if (!inputs?.projectPath) return "Project path is not defined";
  const envListRes = await envUtil.listEnv(inputs.projectPath, true);
  if (envListRes.isErr()) {
    return getLocalizedString("core.getQuestionNewTargetEnvironmentName.validation4");
  }

  inputs.existingEnvNames = envListRes.value; //cache existing env names

  const found =
    envListRes.value.find(
      (env) => env.localeCompare(targetEnvName, undefined, { sensitivity: "base" }) === 0
    ) !== undefined;
  if (found) {
    return getLocalizedString(
      "core.getQuestionNewTargetEnvironmentName.validation5",
      targetEnvName
    );
  } else {
    return undefined;
  }
}
export function newTargetEnvQuestion(): TextInputQuestion {
  return {
    type: "text",
    name: QuestionNames.NewTargetEnvName,
    cliName: "name",
    cliDescription: "Specifies the new environment name.",
    cliType: "argument",
    title: getLocalizedString("core.getQuestionNewTargetEnvironmentName.title"),
    validation: {
      validFunc: newEnvNameValidation,
    },
    placeholder: getLocalizedString("core.getQuestionNewTargetEnvironmentName.placeholder"),
  };
}
// export const lastUsedMark = " (last used)";
// let lastUsedEnv: string | undefined;
// export function reOrderEnvironments(environments: Array<string>): Array<string> {
//   if (!lastUsedEnv) {
//     return environments;
//   }

//   const index = environments.indexOf(lastUsedEnv);
//   if (index === -1) {
//     return environments;
//   }

//   return [lastUsedEnv + lastUsedMark]
//     .concat(environments.slice(0, index))
//     .concat(environments.slice(index + 1));
// }
export function selectSourceEnvQuestion(): SingleSelectQuestion {
  return {
    type: "singleSelect",
    name: QuestionNames.SourceEnvName,
    cliName: "env",
    title: getLocalizedString("core.QuestionSelectSourceEnvironment.title"),
    cliDescription: "Specifies an existing environment name to copy from.",
    staticOptions: [],
    dynamicOptions: async (inputs: Inputs) => {
      if (inputs.existingEnvNames) {
        const envList = inputs.existingEnvNames;
        return envList;
      } else if (inputs.projectPath) {
        const envListRes = await envUtil.listEnv(inputs.projectPath, true);
        if (envListRes.isErr()) {
          throw envListRes.error;
        }
        return envListRes.value;
      }
      return [];
    },
    skipSingleOption: true,
    forgetLastValue: true,
  };
}

export function createNewEnvQuestionNode(): IQTreeNode {
  return {
    data: newTargetEnvQuestion(),
    children: [
      {
        data: selectSourceEnvQuestion(),
      },
    ],
  };
}

// add Plugin to a declarative Copilot project
export function addPluginQuestionNode(): IQTreeNode {
  return {
    data: apiPluginStartQuestion(true),
    children: [
      {
        data: pluginManifestQuestion(),
        condition: {
          equals: ApiPluginStartOptions.existingPlugin().id,
        },
      },
      {
        data: pluginApiSpecQuestion(),
        condition: {
          equals: ApiPluginStartOptions.existingPlugin().id,
        },
      },
      {
        data: apiSpecLocationQuestion(),
        condition: (inputs: Inputs) => {
          return (
            !featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
            inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id
          );
        },
      },
      {
        data: apiOperationQuestion(true, true),
        condition: (inputs: Inputs) => {
          return (
            !featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
            inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id
          );
        },
      },
      {
        data: selectTeamsAppManifestQuestion(),
      },
    ],
  };
}

// add Knowledge to a declarative Copilot project
export function addKnowledgeQuestionNode(): IQTreeNode {
  return {
    data: addKnowledgeStartQuestion(true),
    children: [],
  };
}

export function kiotaRegenerateQuestion(): IQTreeNode {
  return {
    data: selectTeamsAppManifestQuestion(),
  };
}

export function addAuthActionQuestion(): IQTreeNode {
  return {
    data: pluginManifestQuestion(),
    children: [
      {
        data: apiSpecFromPluginManifestQuestion(),
        condition: async (inputs: Inputs) => {
          const pluginManifestPath = inputs[QuestionNames.PluginManifestFilePath];
          if (!!!pluginManifestPath) {
            return false;
          }
          const pluginManifest = (await fs.readJson(
            pluginManifestPath as string
          )) as PluginManifestSchema;
          const specs = pluginManifest
            .runtimes!.filter((runtime) => runtime.type === "OpenApi")
            .map((runtime) => runtime.spec.url);
          const spesDedup = [...new Set(specs)];
          if (spesDedup.length === 1) {
            inputs[QuestionNames.ApiSpecLocation] = spesDedup[0];
            return false;
          }
          return true;
        },
      },
      {
        data: apiFromPluginManifestQuestion(),
        condition: async (inputs: Inputs) => {
          const pluginManifestPath = inputs[QuestionNames.PluginManifestFilePath];
          const apiSpecPath = inputs[QuestionNames.ApiSpecLocation];
          if (!!!pluginManifestPath || !!!apiSpecPath) {
            return false;
          }
          const pluginManifest = (await fs.readJson(
            pluginManifestPath as string
          )) as PluginManifestSchema;
          const apis: string[] = [];
          pluginManifest
            .runtimes!.filter(
              (runtime) => runtime.type === "OpenApi" && runtime.spec.url === apiSpecPath
            )
            .forEach((runtime) => {
              apis.push(...(runtime.run_for_functions as string[]));
            });
          const apisDedup = [...new Set(apis)];
          if (apisDedup.length === 1) {
            inputs[QuestionNames.ApiOperation] = apisDedup;
            return false;
          }
          return true;
        },
      },
      {
        data: authNameQuestion(),
      },
      {
        data: addAuthActionAuthTypeQuestion(),
      },
      oauthParametersQuestion(),
      apiKeyParameterQuestion(),
      microsoftEntraParameterQuestion(),
    ],
  };
}

export function addAuthActionAuthTypeQuestion(): SingleSelectQuestion {
  return {
    type: "singleSelect",
    name: QuestionNames.ApiAuth,
    title: getLocalizedString("core.createProjectQuestion.apiMessageExtensionAuth.title"),
    placeholder: getLocalizedString(
      "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
    ),
    cliDescription: "The authentication type for the API.",
    staticOptions: AddAuthActionAuthTypeOptions.all(),
    default: AddAuthActionAuthTypeOptions.bearerToken().id,
  };
}

export function oauthParametersQuestion(): IQTreeNode {
  return {
    data: oauthAuthorizationUrlQuestion(),
    condition: (inputs: Inputs) => {
      return inputs[QuestionNames.ApiAuth] === AddAuthActionAuthTypeOptions.oauth().id;
    },
    children: [
      {
        data: oauthTokenUrlQuestion(),
      },
      {
        data: oauthRefreshUrlQuestion(),
      },
      {
        data: oauthScopeQuestion(),
      },
      {
        data: oauthPKCEQuestion(),
      },
    ],
  };
}

export function oauthAuthorizationUrlQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.OAuthAuthorizationUrl,
    title: getLocalizedString("core.addAuthActionQuestion.OAuthAuthorizationUrl.title"),
    type: "text",
    cliDescription: "Authorization Url for oauth.",
  };
}

export function oauthTokenUrlQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.OAuthTokenUrl,
    title: getLocalizedString("core.addAuthActionQuestion.OAuthTokenUrl.title"),
    type: "text",
    cliDescription: "Token Url for oauth.",
  };
}

export function oauthRefreshUrlQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.OAuthRefreshUrl,
    title: getLocalizedString("core.addAuthActionQuestion.OAuthRefreshUrl.title"),
    type: "text",
    cliDescription: "Refresh Url for oauth. Leave it emplt if not needed.",
  };
}

export function oauthScopeQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.OAuthScope,
    title: getLocalizedString("core.addAuthActionQuestion.OAuthScope.title"),
    type: "text",
    cliDescription: "Scope for oauth.",
  };
}

export function oauthPKCEQuestion(): SingleSelectQuestion {
  return {
    name: QuestionNames.OauthPKCE,
    title: getLocalizedString("core.addAuthActionQuestion.OauthPKCE.title"),
    type: "singleSelect",
    staticOptions: [
      {
        id: "true",
        label: getLocalizedString("core.addAuthActionQuestion.OauthPKCE.true"),
      },
      {
        id: "false",
        label: getLocalizedString("core.addAuthActionQuestion.OauthPKCE.false"),
      },
    ],
    default: "false",
  };
}

export function apiKeyParameterQuestion(): IQTreeNode {
  return {
    data: apiKeyInQuestion(),
    condition: (inputs: Inputs) => {
      return inputs[QuestionNames.ApiAuth] === AddAuthActionAuthTypeOptions.apiKey().id;
    },
    children: [
      {
        data: apiKeyNameQuestion(),
      },
    ],
  };
}

export function apiKeyInQuestion(): SingleSelectQuestion {
  return {
    name: QuestionNames.ApiKeyIn,
    title: getLocalizedString("core.addAuthActionQuestion.ApiKeyIn.title"),
    type: "singleSelect",
    staticOptions: [
      {
        id: "header",
        label: getLocalizedString("core.addAuthActionQuestion.ApiKeyIn.header"),
      },
      {
        id: "query",
        label: getLocalizedString("core.addAuthActionQuestion.ApiKeyIn.query"),
      },
    ],
    default: "header",
  };
}

export function apiKeyNameQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.ApiKeyName,
    title: getLocalizedString("core.addAuthActionQuestion.ApiKeyName.title"),
    type: "text",
    cliDescription: "Name of the API key.",
  };
}

export function microsoftEntraParameterQuestion(): IQTreeNode {
  return {
    data: oauthScopeQuestion(),
    condition: (inputs: Inputs) => {
      return inputs[QuestionNames.ApiAuth] === AddAuthActionAuthTypeOptions.microsoftEntra().id;
    },
  };
}

export function apiSpecFromPluginManifestQuestion(): SingleSelectQuestion {
  return {
    name: QuestionNames.ApiSpecLocation,
    title: getLocalizedString("core.addAuthActionQuestion.ApiSpecLocation.title"),
    placeholder: getLocalizedString("core.addAuthActionQuestion.ApiSpecLocation.placeholder"),
    type: "singleSelect",
    staticOptions: [],
    cliDescription: "OpenAPI specification to add Auth configuration.",
    dynamicOptions: async (inputs: Inputs) => {
      const pluginManifestPath = inputs[QuestionNames.PluginManifestFilePath];
      const pluginManifest = (await fs.readJson(pluginManifestPath)) as PluginManifestSchema;
      const specs = pluginManifest
        .runtimes!.filter((runtime) => runtime.type === "OpenApi")
        .map((runtime) => runtime.spec.url as string);
      return [...new Set(specs)];
    },
  };
}

export function apiFromPluginManifestQuestion(): MultiSelectQuestion {
  return {
    name: QuestionNames.ApiOperation,
    title: getLocalizedString("core.addAuthActionQuestion.ApiOperation.title"),
    type: "multiSelect",
    staticOptions: [],
    placeholder: getLocalizedString("core.addAuthActionQuestion.ApiOperation.placeholder"),
    cliDescription: "API to add Auth configuration.",
    dynamicOptions: async (inputs: Inputs) => {
      const pluginManifestPath = inputs[QuestionNames.PluginManifestFilePath];
      const apiSpecPath = inputs[QuestionNames.ApiSpecLocation];
      const pluginManifest = (await fs.readJson(pluginManifestPath)) as PluginManifestSchema;
      const apis: string[] = [];
      pluginManifest
        .runtimes!.filter(
          (runtime) => runtime.type === "OpenApi" && runtime.spec.url === apiSpecPath
        )
        .forEach((runtime) => {
          apis.push(...(runtime.run_for_functions as string[]));
        });
      return [...new Set(apis)];
    },
  };
}

export function authNameQuestion(): TextInputQuestion {
  return {
    name: QuestionNames.AuthName,
    title: getLocalizedString("core.addAuthActionQuestion.authName.title"),
    type: "text",
    cliDescription: "Name of Auth Configuration.",
    additionalValidationOnAccept: {
      validFunc: (input: string, inputs?: Inputs): string | undefined => {
        if (!inputs) {
          throw new Error("inputs is undefined"); // should never happen
        }
        inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.newApi().id;
        return;
      },
    },
  };
}

export function apiSpecApiKeyConfirmQestion(): ConfirmQuestion {
  return {
    name: QuestionNames.ApiSpecApiKeyConfirm,
    title: getLocalizedString("core.createProjectQuestion.ApiKeyConfirm"),
    type: "confirm",
    default: true,
  };
}

export function apiSpecApiKeyQuestion(): IQTreeNode {
  return {
    data: {
      type: "text",
      name: QuestionNames.ApiSpecApiKey,
      cliShortName: "k",
      password: true,
      title: getLocalizedString("core.createProjectQuestion.ApiKey"),
      cliDescription: "Api key for OpenAPI spec.",
      forgetLastValue: true,
      validation: {
        validFunc: (input: string): string | undefined => {
          if (input.length < 10 || input.length > 512) {
            return getLocalizedString("core.createProjectQuestion.invalidApiKey.message");
          }

          return undefined;
        },
      },
      additionalValidationOnAccept: {
        validFunc: (input: string, inputs?: Inputs): string | undefined => {
          if (!inputs) {
            throw new Error("inputs is undefined"); // should never happen
          }

          process.env[QuestionNames.ApiSpecApiKey] = input;
          return;
        },
      },
    },
    condition: (inputs: Inputs) => {
      return (
        inputs.outputEnvVarNames &&
        !process.env[inputs.outputEnvVarNames.get("registrationId")] &&
        !inputs.primaryClientSecret &&
        !inputs.secondaryClientSecret
      );
    },
    children: [
      {
        data: apiSpecApiKeyConfirmQestion(),
      },
    ],
  };
}

export function oauthQuestion(): IQTreeNode {
  return {
    data: { type: "group" },
    condition: (inputs: Inputs) => {
      return (
        inputs.outputEnvVarNames && !process.env[inputs.outputEnvVarNames.get("configurationId")]
      );
    },
    children: [
      {
        data: oauthClientIdQuestion(),
        condition: (inputs: Inputs) => {
          return !inputs.clientId;
        },
      },
      {
        data: oauthClientSecretQuestion(),
        condition: (inputs: Inputs) => {
          return (
            !inputs.isPKCEEnabled &&
            !inputs.clientSecret &&
            (!inputs.identityProvider || inputs.identityProvider === "Custom")
          );
        },
      },
      {
        data: oauthConfirmQestion(),
        condition: (inputs: Inputs) => {
          return (
            !inputs.isPKCEEnabled &&
            (!inputs.clientSecret || !inputs.clientId) &&
            (!inputs.identityProvider || inputs.identityProvider === "Custom")
          );
        },
      },
    ],
  };
}

export function uninstallQuestionNode(): IQTreeNode {
  return {
    data: {
      type: "group",
    },
    children: [
      {
        data: uninstallModeQuestion(),
        condition: () => {
          return true;
        },
        children: [
          {
            data: {
              type: "text",
              name: QuestionNames.ManifestId,
              title: getLocalizedString("core.uninstallQuestion.manifestId"),
            },
            condition: (input: UninstallInputs) => {
              return input[QuestionNames.UninstallMode] === QuestionNames.UninstallModeManifestId;
            },
          },
          {
            data: {
              type: "text",
              name: QuestionNames.Env,
              title: getLocalizedString("core.uninstallQuestion.env"),
            },
            condition: (input: UninstallInputs) => {
              return input[QuestionNames.UninstallMode] === QuestionNames.UninstallModeEnv;
            },
            children: [
              {
                data: uninstallProjectPathQuestion(),
                condition: () => {
                  return true;
                },
              },
            ],
          },
          {
            data: uninstallOptionQuestion(),
            condition: (input: UninstallInputs) => {
              return (
                input[QuestionNames.UninstallMode] === QuestionNames.UninstallModeManifestId ||
                input[QuestionNames.UninstallMode] === QuestionNames.UninstallModeEnv
              );
            },
          },
          {
            data: {
              type: "text",
              name: QuestionNames.TitleId,
              title: getLocalizedString("core.uninstallQuestion.titleId"),
            },
            condition: (input: UninstallInputs) => {
              return input[QuestionNames.UninstallMode] === QuestionNames.UninstallModeTitleId;
            },
          },
        ],
      },
    ],
  };
}

function uninstallModeQuestion(): SingleSelectQuestion {
  return {
    name: QuestionNames.UninstallMode,
    title: getLocalizedString("core.uninstallQuestion.chooseMode"),
    type: "singleSelect",
    staticOptions: [
      {
        id: QuestionNames.UninstallModeManifestId,
        label: getLocalizedString("core.uninstallQuestion.manifestIdMode"),
        detail: getLocalizedString("core.uninstallQuestion.manifestIdMode.detail"),
      },
      {
        id: QuestionNames.UninstallModeEnv,
        label: getLocalizedString("core.uninstallQuestion.envMode"),
        detail: getLocalizedString("core.uninstallQuestion.envMode.detail"),
      },
      {
        id: QuestionNames.UninstallModeTitleId,
        label: getLocalizedString("core.uninstallQuestion.titleIdMode"),
        detail: getLocalizedString("core.uninstallQuestion.titleIdMode.detail"),
      },
    ],
    default: QuestionNames.UninstallModeManifestId,
  };
}

function uninstallOptionQuestion(): MultiSelectQuestion {
  return {
    name: QuestionNames.UninstallOptions,
    title: getLocalizedString("core.uninstallQuestion.chooseOption"),
    type: "multiSelect",
    staticOptions: [
      {
        id: QuestionNames.UninstallOptionM365,
        label: getLocalizedString("core.uninstallQuestion.m365Option"),
      },
      {
        id: QuestionNames.UninstallOptionTDP,
        label: getLocalizedString("core.uninstallQuestion.tdpOption"),
      },
      {
        id: QuestionNames.UninstallOptionBot,
        label: getLocalizedString("core.uninstallQuestion.botOption"),
      },
    ],
  };
}
function uninstallProjectPathQuestion(): FolderQuestion {
  return {
    type: "folder",
    name: QuestionNames.ProjectPath,
    title: getLocalizedString("core.uninstallQuestion.projectPath"),
    cliDescription: "Project Path for uninstall",
    placeholder: "./",
    default: "./",
  };
}

function oauthClientIdQuestion(): TextInputQuestion {
  return {
    type: "text",
    name: QuestionNames.OauthClientId,
    cliShortName: "i",
    title: getLocalizedString("core.createProjectQuestion.OauthClientId"),
    cliDescription: "Oauth client id for OpenAPI spec.",
    forgetLastValue: true,
    additionalValidationOnAccept: {
      validFunc: (input: string, inputs?: Inputs): string | undefined => {
        if (!inputs) {
          throw new Error("inputs is undefined"); // should never happen
        }

        process.env[QuestionNames.OauthClientId] = input;
        return;
      },
    },
  };
}

function oauthConfirmQestion(): ConfirmQuestion {
  return {
    name: QuestionNames.OauthConfirm,
    title: getLocalizedString("core.createProjectQuestion.OauthClientSecretConfirm"),
    type: "confirm",
    default: true,
  };
}

function oauthClientSecretQuestion(): TextInputQuestion {
  return {
    type: "text",
    name: QuestionNames.OauthClientSecret,
    cliShortName: "c",
    password: true,
    title: getLocalizedString("core.createProjectQuestion.OauthClientSecret"),
    cliDescription: "Oauth client secret for OpenAPI spec.",
    forgetLastValue: true,
    validation: {
      validFunc: (input: string): string | undefined => {
        if (input.length < 10 || input.length > 512) {
          return getLocalizedString("core.createProjectQuestion.invalidApiKey.message");
        }

        return undefined;
      },
    },
    additionalValidationOnAccept: {
      validFunc: (input: string, inputs?: Inputs): string | undefined => {
        if (!inputs) {
          throw new Error("inputs is undefined"); // should never happen
        }

        process.env[QuestionNames.OauthClientSecret] = input;
        return;
      },
    },
  };
}

export function syncManifestQuestionNode(): IQTreeNode {
  return {
    data: {
      type: "group",
    },
    children: [
      {
        data: {
          type: "folder",
          name: QuestionNames.ProjectPath,
          title: getLocalizedString("core.syncManifest.projectPath"),
          cliDescription: "Project Path",
          placeholder: "./",
          default: (inputs: Inputs) =>
            CLIPlatforms.includes(inputs.platform)
              ? "./"
              : path.join(os.homedir(), ConstantString.RootFolder),
        },
      },
      {
        data: {
          type: "text",
          name: QuestionNames.Env,
          title: getLocalizedString("core.syncManifest.env"),
          cliDescription: "Target Teams Toolkit Environment",
        },
      },
      {
        data: {
          type: "text",
          name: QuestionNames.TeamsAppId,
          title: getLocalizedString("core.syncManifest.teamsAppId"),
          cliDescription: "Teams App ID (optional)",
        },
      },
    ],
  };
}
