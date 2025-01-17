// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ApiOperation,
  ConditionFunc,
  Inputs,
  IQTreeNode,
  MultiSelectQuestion,
  OptionItem,
  SingleFileOrInputQuestion,
  StringValidation,
} from "@microsoft/teamsfx-api";
import * as jsonschema from "jsonschema";
import { getLocalizedString } from "../../../common/localizeUtils";
import { DevEnvironmentSetupError } from "../../../component/generator/spfx/error";
import { Constants } from "../../../component/generator/spfx/utils/constants";
import { Utils } from "../../../component/generator/spfx/utils/utils";
import { QuestionNames, SPFxVersionOptionIds } from "../../constants";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  BotCapabilityOptions,
  CustomCopilotRagOptions,
  MeArchitectureOptions,
  MeCapabilityOptions,
  NotificationBotOptions,
  setTemplateName,
  TabCapabilityOptions,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { assembleError, EmptyOptionError } from "../../../error/common";
import { Correlator } from "../../../common/correlator";
import { createContext } from "../../../common/globalVars";
import { listOperations } from "../../../component/generator/apiSpec/helper";
import { isValidHttpUrl } from "../../../common/stringUtils";
import fs from "fs-extra";

export function apiSpecNode(condition: StringValidation | ConditionFunc): IQTreeNode {
  return {
    condition: condition,
    data: { type: "group", name: QuestionNames.FromExistingApi },
    children: [
      {
        data: apiSpecLocationQuestion(),
      },
      {
        condition: (inputs: Inputs) => {
          return !inputs[QuestionNames.ApiPluginManifestPath];
        },
        data: apiOperationQuestion(),
      },
    ],
  };
}

export function apiOperationQuestion(): MultiSelectQuestion {
  let placeholder = "";
  const isPlugin = (inputs?: Inputs): boolean => {
    return !!inputs && inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id;
  };

  return {
    type: "multiSelect",
    name: QuestionNames.ApiOperation,
    title: (inputs: Inputs) => {
      return isPlugin(inputs)
        ? getLocalizedString("core.createProjectQuestion.apiSpec.copilotOperation.title")
        : getLocalizedString("core.createProjectQuestion.apiSpec.operation.title");
    },
    placeholder: (inputs: Inputs) => {
      const isPlugin = inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id;
      if (isPlugin) {
        placeholder = getLocalizedString(
          "core.createProjectQuestion.apiSpec.operation.plugin.placeholder"
        );
      } else {
        placeholder = getLocalizedString(
          "core.createProjectQuestion.apiSpec.operation.apikey.placeholder"
        );
      }
      return placeholder;
    },
    forgetLastValue: true,
    staticOptions: [],
    validation: {
      validFunc: (input: string[], inputs?: Inputs): string | undefined => {
        if (!inputs) {
          throw new Error("inputs is undefined"); // should never happen
        }
        if (
          input.length < 1 ||
          (input.length > 10 &&
            inputs[QuestionNames.CustomCopilotRag] !== CustomCopilotRagOptions.customApi().id &&
            inputs[QuestionNames.ProjectType] !== ProjectTypeOptions.Agent().id)
        ) {
          return getLocalizedString(
            "core.createProjectQuestion.apiSpec.operation.invalidMessage",
            input.length,
            10
          );
        }
        const operations: ApiOperation[] = inputs.supportedApisFromApiSpec as ApiOperation[];

        const authNames: Set<string> = new Set();
        const serverUrls: Set<string> = new Set();
        for (const inputItem of input) {
          const operation = operations.find((op) => op.id === inputItem);
          if (operation) {
            if (operation.data.authName) {
              authNames.add(operation.data.authName);
              serverUrls.add(operation.data.serverUrl);
            }
          }
        }

        if (serverUrls.size > 1) {
          return getLocalizedString(
            "core.createProjectQuestion.apiSpec.operation.multipleServer",
            Array.from(serverUrls).join(", ")
          );
        }

        const seenAuthNames = new Set<string>();
        const uniqueAuthApis = operations.filter((api) => {
          if (
            !!api.data.authName &&
            input.includes(api.id) &&
            !seenAuthNames.has(api.data.authName)
          ) {
            seenAuthNames.add(api.data.authName);
            return true;
          }
          return false;
        });
        inputs.apiAuthData = uniqueAuthApis.map((authApi) => authApi.data);
      },
    },
    dynamicOptions: (inputs: Inputs) => {
      if (!inputs.supportedApisFromApiSpec) {
        throw new EmptyOptionError(QuestionNames.ApiOperation, "question");
      }

      const operations = inputs.supportedApisFromApiSpec as ApiOperation[];

      return operations;
    },
  };
}

const maximumLengthOfDetailsErrorMessageInInputBox = 90;
export function apiSpecLocationQuestion(): SingleFileOrInputQuestion {
  const correlationId = Correlator.getId(); // This is a workaround for VSCode which will lose correlation id when user accepts the value.
  const validationOnAccept = async (
    input: string,
    inputs?: Inputs
  ): Promise<string | undefined> => {
    try {
      if (!inputs) {
        throw new Error("inputs is undefined"); // should never happen
      }
      const context = createContext();
      const res = await listOperations(context, input.trim(), inputs, true, false, correlationId);
      if (res.isOk()) {
        inputs.supportedApisFromApiSpec = res.value;
      } else {
        const errors = res.error;
        if (
          errors.length === 1 &&
          errors[0].content.length <= maximumLengthOfDetailsErrorMessageInInputBox
        ) {
          return errors[0].content;
        } else {
          return getLocalizedString(
            "core.createProjectQuestion.apiSpec.multipleValidationErrors.vscode.message"
          );
        }
      }
    } catch (e) {
      const error = assembleError(e);
      throw error;
    }
  };
  return {
    type: "singleFileOrText",
    name: QuestionNames.ApiSpecLocation,
    cliShortName: "a",
    cliDescription: "OpenAPI description document location.",
    title: getLocalizedString("core.createProjectQuestion.apiSpec.title"),
    forgetLastValue: true,
    inputBoxConfig: {
      type: "innerText",
      title: getLocalizedString("core.createProjectQuestion.apiSpec.title"),
      placeholder: getLocalizedString("core.createProjectQuestion.apiSpec.placeholder"),
      name: "input-api-spec-url",
      step: 2, // Add "back" button
      validation: {
        validFunc: (input: string, inputs?: Inputs): Promise<string | undefined> => {
          const result = isValidHttpUrl(input.trim())
            ? undefined
            : getLocalizedString("core.createProjectQuestion.invalidUrl.message");
          return Promise.resolve(result);
        },
      },
    },
    inputOptionItem: {
      id: "input",
      label: `$(cloud) ` + getLocalizedString("core.createProjectQuestion.apiSpecInputUrl.label"),
    },
    filters: {
      files: ["json", "yml", "yaml"],
    },
    validation: {
      validFunc: async (input: string, inputs?: Inputs): Promise<string | undefined> => {
        if (!isValidHttpUrl(input.trim()) && !(await fs.pathExists(input.trim()))) {
          return "Please enter a valid HTTP URL without authentication to access your OpenAPI description document or enter a file path of your local OpenAPI description document.";
        }

        return await validationOnAccept(input, inputs);
      },
    },
  };
}

export function botProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Bot
    condition: { equals: ProjectTypeOptions.botOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.bot.title"),
      type: "singleSelect",
      staticOptions: [
        BotCapabilityOptions.basicBot(),
        BotCapabilityOptions.notificationBot(),
        BotCapabilityOptions.commandBot(),
        BotCapabilityOptions.workflowBot(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [
      {
        // 2.3.1 Notification bot trigger sub-tree
        condition: { equals: BotCapabilityOptions.notificationBotId },
        data: {
          name: QuestionNames.BotTrigger,
          title: getLocalizedString("plugins.bot.questionHostTypeTrigger.title"),
          type: "singleSelect",
          staticOptions: [
            NotificationBotOptions.appService(),
            NotificationBotOptions.functionsHttpAndTimerTrigger(),
            NotificationBotOptions.functionsHttpTrigger(),
            NotificationBotOptions.functionsTimerTrigger(),
          ],
          placeholder: getLocalizedString("plugins.bot.questionHostTypeTrigger.placeholder"),
          onDidSelection: setTemplateName,
        },
      },
    ],
  };
}

export function tabProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Tab
    condition: { equals: ProjectTypeOptions.tab().id },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.tab.title"),
      type: "singleSelect",
      staticOptions: [
        TabCapabilityOptions.nonSsoTab(),
        TabCapabilityOptions.m365SsoLaunchPage(),
        TabCapabilityOptions.dashboardTab(),
        TabCapabilityOptions.SPFxTab(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [
      {
        //SPFx sub-tree
        condition: { equals: TabCapabilityOptions.SPFxTab().id },
        data: {
          type: "singleSelect",
          name: QuestionNames.SPFxSolution,
          title: getLocalizedString("plugins.spfx.questions.spfxSolution.title"),
          staticOptions: [
            {
              id: "new",
              label: getLocalizedString("plugins.spfx.questions.spfxSolution.createNew"),
              detail: getLocalizedString("plugins.spfx.questions.spfxSolution.createNew.detail"),
            },
            {
              id: "import",
              label: getLocalizedString("plugins.spfx.questions.spfxSolution.importExisting"),
              detail: getLocalizedString(
                "plugins.spfx.questions.spfxSolution.importExisting.detail"
              ),
            },
          ],
          default: "new",
        },
        children: [
          {
            condition: { equals: "new" },
            data: { type: "group" },
            children: [
              {
                data: {
                  type: "singleSelect",
                  name: QuestionNames.SPFxInstallPackage,
                  title: getLocalizedString("plugins.spfx.questions.packageSelect.title"),
                  staticOptions: [],
                  placeholder: getLocalizedString(
                    "plugins.spfx.questions.packageSelect.placeholder"
                  ),
                  dynamicOptions: async (inputs: Inputs): Promise<OptionItem[]> => {
                    const versions = await Promise.all([
                      Utils.findGloballyInstalledVersion(
                        undefined,
                        Constants.GeneratorPackageName,
                        0,
                        false
                      ),
                      Utils.findLatestVersion(undefined, Constants.GeneratorPackageName, 5),
                      Utils.findGloballyInstalledVersion(
                        undefined,
                        Constants.YeomanPackageName,
                        0,
                        false
                      ),
                    ]);

                    inputs.globalSpfxPackageVersion = versions[0];
                    inputs.latestSpfxPackageVersion = versions[1];
                    inputs.globalYeomanPackageVersion = versions[2];

                    return [
                      {
                        id: SPFxVersionOptionIds.installLocally,

                        label:
                          versions[1] !== undefined
                            ? getLocalizedString(
                                "plugins.spfx.questions.packageSelect.installLocally.withVersion.label",
                                "v" + versions[1]
                              )
                            : getLocalizedString(
                                "plugins.spfx.questions.packageSelect.installLocally.noVersion.label"
                              ),
                      },
                      {
                        id: SPFxVersionOptionIds.globalPackage,
                        label:
                          versions[0] !== undefined
                            ? getLocalizedString(
                                "plugins.spfx.questions.packageSelect.useGlobalPackage.withVersion.label",
                                "v" + versions[0]
                              )
                            : getLocalizedString(
                                "plugins.spfx.questions.packageSelect.useGlobalPackage.noVersion.label"
                              ),
                        description: getLocalizedString(
                          "plugins.spfx.questions.packageSelect.useGlobalPackage.detail",
                          Constants.RecommendedLowestSpfxVersion
                        ),
                      },
                    ];
                  },
                  default: SPFxVersionOptionIds.installLocally,
                  validation: {
                    validFunc: (
                      input: string,
                      previousInputs?: Inputs
                    ): Promise<string | undefined> => {
                      if (input === SPFxVersionOptionIds.globalPackage) {
                        const hasPackagesInstalled =
                          !!previousInputs &&
                          !!previousInputs.globalSpfxPackageVersion &&
                          !!previousInputs.globalYeomanPackageVersion;
                        if (!hasPackagesInstalled) {
                          return Promise.reject(DevEnvironmentSetupError());
                        }
                      }
                      return Promise.resolve(undefined);
                    },
                  },
                  isBoolean: true,
                },
              },
              {
                data: {
                  type: "singleSelect",
                  name: QuestionNames.SPFxFramework,
                  title: getLocalizedString("plugins.spfx.questions.framework.title"),
                  staticOptions: [
                    { id: "react", label: "React" },
                    { id: "minimal", label: "Minimal" },
                    { id: "none", label: "None" },
                  ],
                  placeholder: "Select an option",
                  default: "react",
                },
              },
              {
                data: {
                  type: "text",
                  name: QuestionNames.SPFxWebpartName,
                  title: getLocalizedString("plugins.spfx.questions.webpartName"),
                  default: Constants.DEFAULT_WEBPART_NAME,
                  validation: {
                    validFunc: (input: string): string | undefined => {
                      const schema = {
                        pattern: "^[a-zA-Z_][a-zA-Z0-9_]*$",
                      };
                      const validateRes = jsonschema.validate(input, schema);
                      if (validateRes.errors && validateRes.errors.length > 0) {
                        return getLocalizedString(
                          "plugins.spfx.questions.webpartName.error.notMatch",
                          input,
                          schema.pattern
                        );
                      }
                      return undefined;
                    },
                  },
                },
              },
            ],
          },
          {
            condition: { equals: "import" },
            data: {
              type: "folder",
              name: QuestionNames.SPFxFolder,
              title: getLocalizedString("core.spfxFolder.title"),
              placeholder: getLocalizedString("core.spfxFolder.placeholder"),
            },
          },
        ],
      },
    ],
  };
}

export function meProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Messaging Extension
    condition: { equals: ProjectTypeOptions.meOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.messageExtension.title"),
      type: "singleSelect",
      staticOptions: [
        MeCapabilityOptions.m365SearchMe(),
        MeCapabilityOptions.collectFormMe(),
        MeCapabilityOptions.linkUnfurling(),
      ],
      placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [
      {
        // Search ME sub-tree
        condition: { equals: MeCapabilityOptions.m365SearchMe().id },
        data: {
          name: QuestionNames.MeArchitectureType,
          title: getLocalizedString("core.createProjectQuestion.meArchitecture.title"),
          type: "singleSelect",
          staticOptions: [
            MeArchitectureOptions.newApi(),
            MeArchitectureOptions.apiSpec(),
            MeArchitectureOptions.botPlugin(),
          ],
          default: MeArchitectureOptions.newApi().id,
          placeholder: getLocalizedString(
            "core.createProjectQuestion.projectType.copilotExtension.placeholder"
          ),
          forgetLastValue: true,
          skipSingleOption: true,
          onDidSelection: setTemplateName,
        },
        children: [
          {
            condition: { equals: MeArchitectureOptions.newApi().id },
            data: {
              type: "singleSelect",
              name: QuestionNames.ApiAuth,
              title: getLocalizedString("core.createProjectQuestion.apiMessageExtensionAuth.title"),
              placeholder: getLocalizedString(
                "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
              ),
              staticOptions: [
                ApiAuthOptions.none(),
                ApiAuthOptions.bearerToken(),
                ApiAuthOptions.microsoftEntra(),
              ],
              default: ApiAuthOptions.none().id,
              onDidSelection: setTemplateName,
            },
          },
          apiSpecNode({ equals: MeArchitectureOptions.apiSpec().id }),
        ],
      },
    ],
  };
}
