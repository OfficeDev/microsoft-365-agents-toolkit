// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ApiOperation,
  Inputs,
  IQTreeNode,
  MultiSelectQuestion,
  OptionItem,
  SingleFileOrInputQuestion,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as jsonschema from "jsonschema";
import * as os from "os";
import * as path from "path";
import { ConstantString } from "../../common/constants";
import { Correlator } from "../../common/correlator";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";
import { createContext } from "../../common/globalVars";
import { getLocalizedString } from "../../common/localizeUtils";
import { isValidHttpUrl } from "../../common/stringUtils";
import { listOperations } from "../../component/generator/apiSpec/helper";
import { Constants } from "../../component/generator/spfx/utils/constants";
import { assembleError, EmptyOptionError, FileNotFoundError } from "../../error/common";
import {
  AppNamePattern,
  HostType,
  HostTypeTriggerOptionItem,
  NotificationTriggers,
  ProgrammingLanguage,
  QuestionNames,
} from "../constants";
import { TemplateNames, Templates } from "../templates";

function onDidSelectionCapability(selected: string | OptionItem, inputs: Inputs): void {
  if ((selected as OptionItem).data) {
    inputs[QuestionNames.TemplateName] = (selected as OptionItem).data as string;
  }
}

/**
 *
 * FxCore API for scaffold: scaffold(questionModel: IQTreeNode, generators: DefaultTemplateGenerator[]): Promise<Result<any, FxError>>
 *
 */

export function scaffoldQuestionForVSCode(): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      // category tree
      {
        data: {
          name: QuestionNames.ProjectType,
          title: getLocalizedString("core.createProjectQuestion.title"),
          type: "singleSelect",
          staticOptions: [
            ProjectTypeOptions.Agent(),
            // ProjectTypeOptions.customCopilot(),
            ProjectTypeOptions.bot(),
            // ProjectTypeOptions.tab(),
            // ProjectTypeOptions.me(),
            // featureFlagManager.getBooleanValue(FeatureFlags.OfficeMetaOS)
            //   ? ProjectTypeOptions.officeMetaOS()
            //   : featureFlagManager.getBooleanValue(FeatureFlags.OfficeAddin)
            //   ? ProjectTypeOptions.officeAddin()
            //   : ProjectTypeOptions.outlookAddin(),
          ],
        },
        children: [
          {
            // 2.1 Agent sub tree
            condition: { equals: ProjectTypeOptions.copilotAgentOptionId },
            data: {
              name: QuestionNames.Capabilities,
              title: getLocalizedString(
                "core.createProjectQuestion.projectType.copilotExtension.title"
              ),
              type: "singleSelect",
              staticOptions: [CapabilityOptions.apiPlugin(), CapabilityOptions.declarativeAgent()],
              placeholder: getLocalizedString(
                "core.createProjectQuestion.projectType.copilotExtension.placeholder"
              ),
            },
            children: [
              {
                // 2.1.2 declarativeAgent
                condition: { equals: CapabilityOptions.declarativeAgent().id },
                data: {
                  type: "singleSelect",
                  name: QuestionNames.WithPlugin,
                  title: getLocalizedString("core.createProjectQuestion.declarativeCopilot.title"),
                  placeholder: getLocalizedString(
                    "core.createProjectQuestion.declarativeCopilot.placeholder"
                  ),
                  staticOptions: [
                    DeclarativeCopilotTypeOptions.noPlugin(),
                    DeclarativeCopilotTypeOptions.withPlugin(),
                  ],
                  default: DeclarativeCopilotTypeOptions.noPlugin().id,
                },
                children: [
                  {
                    condition: { equals: DeclarativeCopilotTypeOptions.withPlugin().id },
                    data: {
                      type: "singleSelect",
                      name: QuestionNames.ApiPluginType,
                      title: getLocalizedString("core.createProjectQuestion.createApiPlugin.title"),
                      placeholder: getLocalizedString(
                        "core.createProjectQuestion.projectType.copilotExtension.placeholder"
                      ),
                      staticOptions: [
                        ApiPluginStartOptions.newApi(),
                        ApiPluginStartOptions.apiSpec(),
                        // ApiPluginStartOptions.existingPlugin(),
                      ],
                      default: ApiPluginStartOptions.newApi().id,
                      onDidSelection: onDidSelectionCapability,
                    },
                    children: [
                      {
                        condition: { equals: ApiPluginStartOptions.newApi().id },
                        data: {
                          type: "singleSelect",
                          name: QuestionNames.ApiAuth,
                          title: getLocalizedString(
                            "core.createProjectQuestion.apiMessageExtensionAuth.title"
                          ),
                          placeholder: getLocalizedString(
                            "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
                          ),
                          staticOptions: [
                            ApiAuthOptions.none(),
                            ApiAuthOptions.apiKey(),
                            ApiAuthOptions.oauth(),
                          ],
                          default: ApiAuthOptions.none().id,
                        },
                      },
                      {
                        condition: { equals: ApiPluginStartOptions.apiSpec().id },
                        data: { type: "group", name: QuestionNames.FromExistingApi },
                        children: [
                          {
                            data: apiSpecLocationQuestion(),
                          },
                          {
                            data: apiOperationQuestion(),
                            condition: (inputs: Inputs) => {
                              return !inputs[QuestionNames.ApiPluginManifestPath];
                            },
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
          {
            // 2.3 Bot sub tree
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
              onDidSelection: onDidSelectionCapability,
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
                  placeholder: getLocalizedString(
                    "plugins.bot.questionHostTypeTrigger.placeholder"
                  ),
                  onDidSelection: onDidSelectionCapability,
                },
              },
            ],
          },
          // {
          //   // 2.2 customCopilots sub tree
          //   condition: { equals: ProjectTypeOptions.customCopilot().id },
          //   data: {
          //     name: QuestionNames.Capabilities,
          //     title: getLocalizedString(
          //       "core.createProjectQuestion.projectType.customCopilot.title"
          //     ),
          //     type: "singleSelect",
          //     staticOptions: [
          //       CapabilityOptions.customCopilotBasic(),
          //       CapabilityOptions.customCopilotRag(),
          //       CapabilityOptions.customCopilotAssistant(),
          //     ],
          //     placeholder: getLocalizedString(
          //       "core.createProjectQuestion.projectType.customCopilot.placeholder"
          //     ),
          //   },
          //   children: [
          //     {
          //       condition: { equals: CapabilityOptions.customCopilotRag().id },
          //       data: {
          //         type: "singleSelect",
          //         name: QuestionNames.CustomCopilotRag,
          //         title: getLocalizedString(
          //           "core.createProjectQuestion.capability.customCopilotRag.title"
          //         ),
          //         placeholder: getLocalizedString(
          //           "core.createProjectQuestion.capability.customCopilotRag.placeholder"
          //         ),
          //         staticOptions: [
          //           CustomCopilotRagOptions.customize(),
          //           CustomCopilotRagOptions.azureAISearch(),
          //           CustomCopilotRagOptions.customApi(),
          //           CustomCopilotRagOptions.microsoft365(),
          //         ],
          //         default: CustomCopilotRagOptions.customize().id,
          //       },
          //     },
          //     {
          //       condition: { equals: CapabilityOptions.customCopilotAssistant().id },
          //       data: {
          //         type: "singleSelect",
          //         name: QuestionNames.CustomCopilotAssistant,
          //         title: getLocalizedString(
          //           "core.createProjectQuestion.capability.customCopilotAssistant.title"
          //         ),
          //         placeholder: getLocalizedString(
          //           "core.createProjectQuestion.capability.customCopilotAssistant.placeholder"
          //         ),
          //         staticOptions: [
          //           CustomCopilotAssistantOptions.new(),
          //           CustomCopilotAssistantOptions.assistantsApi(),
          //         ],
          //         default: CustomCopilotAssistantOptions.new().id,
          //       },
          //     },
          //     {
          //       data: {
          //         type: "singleSelect",
          //         name: QuestionNames.LLMService,
          //         title: getLocalizedString("core.createProjectQuestion.llmService.title"),
          //         placeholder: getLocalizedString(
          //           "core.createProjectQuestion.llmService.placeholder"
          //         ),
          //         staticOptions: [
          //           {
          //             id: "llm-service-azure-openai",
          //             label: getLocalizedString(
          //               "core.createProjectQuestion.llmServiceAzureOpenAIOption.label"
          //             ),
          //             detail: getLocalizedString(
          //               "core.createProjectQuestion.llmServiceAzureOpenAIOption.detail"
          //             ),
          //           },
          //           {
          //             id: "llm-service-openai",
          //             label: getLocalizedString(
          //               "core.createProjectQuestion.llmServiceOpenAIOption.label"
          //             ),
          //             detail: getLocalizedString(
          //               "core.createProjectQuestion.llmServiceOpenAIOption.detail"
          //             ),
          //           },
          //         ],
          //         skipSingleOption: true,
          //         default: "llm-service-azure-openai",
          //       },
          //       children: [
          //         {
          //           condition: { equals: "llm-service-azure-openai" },
          //           data: {
          //             type: "text",
          //             password: true,
          //             name: QuestionNames.AzureOpenAIKey,
          //             title: getLocalizedString(
          //               "core.createProjectQuestion.llmService.azureOpenAIKey.title"
          //             ),
          //             placeholder: getLocalizedString(
          //               "core.createProjectQuestion.llmService.azureOpenAIKey.placeholder"
          //             ),
          //           },
          //           children: [
          //             {
          //               condition: (inputs: Inputs) => {
          //                 return inputs[QuestionNames.AzureOpenAIKey]?.length > 0;
          //               },
          //               data: {
          //                 type: "text",
          //                 name: QuestionNames.AzureOpenAIEndpoint,
          //                 title: getLocalizedString(
          //                   "core.createProjectQuestion.llmService.azureOpenAIEndpoint.title"
          //                 ),
          //                 placeholder: getLocalizedString(
          //                   "core.createProjectQuestion.llmService.azureOpenAIEndpoint.placeholder"
          //                 ),
          //               },
          //               children: [
          //                 {
          //                   condition: (inputs: Inputs) => {
          //                     return inputs[QuestionNames.AzureOpenAIEndpoint]?.length > 0;
          //                   },
          //                   data: {
          //                     type: "text",
          //                     name: QuestionNames.AzureOpenAIDeploymentName,
          //                     title: getLocalizedString(
          //                       "core.createProjectQuestion.llmService.azureOpenAIDeploymentName.title"
          //                     ),
          //                     placeholder: getLocalizedString(
          //                       "core.createProjectQuestion.llmService.azureOpenAIDeploymentName.placeholder"
          //                     ),
          //                   },
          //                 },
          //               ],
          //             },
          //           ],
          //         },
          //         {
          //           condition: { equals: "llm-service-openai" },
          //           data: {
          //             type: "text",
          //             password: true,
          //             name: QuestionNames.OpenAIKey,
          //             title: getLocalizedString(
          //               "core.createProjectQuestion.llmService.openAIKey.title"
          //             ),
          //             placeholder: getLocalizedString(
          //               "core.createProjectQuestion.llmService.openAIKey.placeholder"
          //             ),
          //           },
          //         },
          //       ],
          //     },
          //   ],
          // },
          // {
          //   // 2.4 Tab sub tree
          //   condition: { equals: ProjectTypeOptions.tab().id },
          //   data: {
          //     name: QuestionNames.Capabilities,
          //     title: getLocalizedString("core.createProjectQuestion.projectType.tab.title"),
          //     type: "singleSelect",
          //     staticOptions: [
          //       CapabilityOptions.nonSsoTab(),
          //       CapabilityOptions.m365SsoLaunchPage(),
          //       CapabilityOptions.dashboardTab(),
          //       CapabilityOptions.SPFxTab(),
          //     ],
          //     placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
          //     onDidSelection: onDidSelectionCapability,
          //   },
          //   children: [
          //     {
          //       // 2.4.1 SPFx sub-tree
          //       // TODO
          //       condition: { equals: CapabilityOptions.SPFxTab().id },
          //       data: {
          //         type: "singleSelect",
          //         name: QuestionNames.SPFxSolution,
          //         title: getLocalizedString("plugins.spfx.questions.spfxSolution.title"),
          //         staticOptions: [
          //           {
          //             id: "new",
          //             label: getLocalizedString("plugins.spfx.questions.spfxSolution.createNew"),
          //             detail: getLocalizedString(
          //               "plugins.spfx.questions.spfxSolution.createNew.detail"
          //             ),
          //           },
          //           {
          //             id: "import",
          //             label: getLocalizedString(
          //               "plugins.spfx.questions.spfxSolution.importExisting"
          //             ),
          //             detail: getLocalizedString(
          //               "plugins.spfx.questions.spfxSolution.importExisting.detail"
          //             ),
          //           },
          //         ],
          //         default: "new",
          //       },
          //       children: [
          //         {
          //           data: { type: "group" },
          //           children: [
          //             {
          //               data: {
          //                 type: "singleSelect",
          //                 name: QuestionNames.SPFxInstallPackage,
          //                 title: getLocalizedString("plugins.spfx.questions.packageSelect.title"),
          //                 staticOptions: [],
          //                 placeholder: getLocalizedString(
          //                   "plugins.spfx.questions.packageSelect.placeholder"
          //                 ),
          //                 dynamicOptions: async (inputs: Inputs): Promise<OptionItem[]> => {
          //                   const versions = await Promise.all([
          //                     Utils.findGloballyInstalledVersion(
          //                       undefined,
          //                       Constants.GeneratorPackageName,
          //                       0,
          //                       false
          //                     ),
          //                     Utils.findLatestVersion(undefined, Constants.GeneratorPackageName, 5),
          //                     Utils.findGloballyInstalledVersion(
          //                       undefined,
          //                       Constants.YeomanPackageName,
          //                       0,
          //                       false
          //                     ),
          //                   ]);

          //                   inputs.globalSpfxPackageVersion = versions[0];
          //                   inputs.latestSpfxPackageVersion = versions[1];
          //                   inputs.globalYeomanPackageVersion = versions[2];

          //                   return [
          //                     {
          //                       id: SPFxVersionOptionIds.installLocally,

          //                       label:
          //                         versions[1] !== undefined
          //                           ? getLocalizedString(
          //                               "plugins.spfx.questions.packageSelect.installLocally.withVersion.label",
          //                               "v" + versions[1]
          //                             )
          //                           : getLocalizedString(
          //                               "plugins.spfx.questions.packageSelect.installLocally.noVersion.label"
          //                             ),
          //                     },
          //                     {
          //                       id: SPFxVersionOptionIds.globalPackage,
          //                       label:
          //                         versions[0] !== undefined
          //                           ? getLocalizedString(
          //                               "plugins.spfx.questions.packageSelect.useGlobalPackage.withVersion.label",
          //                               "v" + versions[0]
          //                             )
          //                           : getLocalizedString(
          //                               "plugins.spfx.questions.packageSelect.useGlobalPackage.noVersion.label"
          //                             ),
          //                       description: getLocalizedString(
          //                         "plugins.spfx.questions.packageSelect.useGlobalPackage.detail",
          //                         Constants.RecommendedLowestSpfxVersion
          //                       ),
          //                     },
          //                   ];
          //                 },
          //                 default: SPFxVersionOptionIds.installLocally,
          //                 validation: {
          //                   validFunc: (
          //                     input: string,
          //                     previousInputs?: Inputs
          //                   ): Promise<string | undefined> => {
          //                     if (input === SPFxVersionOptionIds.globalPackage) {
          //                       const hasPackagesInstalled =
          //                         !!previousInputs &&
          //                         !!previousInputs.globalSpfxPackageVersion &&
          //                         !!previousInputs.globalYeomanPackageVersion;
          //                       if (!hasPackagesInstalled) {
          //                         return Promise.reject(DevEnvironmentSetupError());
          //                       }
          //                     }
          //                     return Promise.resolve(undefined);
          //                   },
          //                 },
          //                 isBoolean: true,
          //               },
          //             },
          //             {
          //               data: {
          //                 type: "singleSelect",
          //                 name: QuestionNames.SPFxFramework,
          //                 cliShortName: "k",
          //                 cliDescription: "Framework.",
          //                 title: getLocalizedString("plugins.spfx.questions.framework.title"),
          //                 staticOptions: [
          //                   { id: "react", label: "React" },
          //                   { id: "minimal", label: "Minimal" },
          //                   { id: "none", label: "None" },
          //                 ],
          //                 placeholder: "Select an option",
          //                 default: "react",
          //               },
          //             },
          //             {
          //               data: {
          //                 type: "text",
          //                 name: QuestionNames.SPFxWebpartName,
          //                 title: getLocalizedString("plugins.spfx.questions.webpartName"),
          //                 default: Constants.DEFAULT_WEBPART_NAME,
          //                 validation: {
          //                   validFunc: (input: string): string | undefined => {
          //                     const schema = {
          //                       pattern: "^[a-zA-Z_][a-zA-Z0-9_]*$",
          //                     };
          //                     const validateRes = jsonschema.validate(input, schema);
          //                     if (validateRes.errors && validateRes.errors.length > 0) {
          //                       return getLocalizedString(
          //                         "plugins.spfx.questions.webpartName.error.notMatch",
          //                         input,
          //                         schema.pattern
          //                       );
          //                     }
          //                     return undefined;
          //                   },
          //                 },
          //               },
          //             },
          //           ],
          //           condition: { equals: "new" },
          //         },
          //         {
          //           data: {
          //             type: "folder",
          //             name: QuestionNames.SPFxFolder,
          //             title: getLocalizedString("core.spfxFolder.title"),
          //             cliDescription:
          //               "Directory or Path that contains the existing SharePoint Framework solution.",
          //             placeholder: getLocalizedString("core.spfxFolder.placeholder"),
          //           },
          //           condition: { equals: "import" },
          //         },
          //       ],
          //     },
          //   ],
          // },
          // {
          //   // 2.5 Messaging Extension sub tree
          //   condition: { equals: ProjectTypeOptions.me().id },
          //   data: {
          //     name: QuestionNames.Capabilities,
          //     title: getLocalizedString(
          //       "core.createProjectQuestion.projectType.messageExtension.title"
          //     ),
          //     type: "singleSelect",
          //     staticOptions: [
          //       CapabilityOptions.m365SearchMe(),
          //       CapabilityOptions.collectFormMe(),
          //       CapabilityOptions.linkUnfurling(),
          //     ],
          //     placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
          //     onDidSelection: onDidSelectionCapability,
          //   },
          //   children: [
          //     {
          //       // 2.5.1 Search ME sub-tree
          //       condition: { equals: CapabilityOptions.m365SearchMe().id },
          //       data: {
          //         name: QuestionNames.MeArchitectureType,
          //         title: getLocalizedString("core.createProjectQuestion.meArchitecture.title"),
          //         type: "singleSelect",
          //         staticOptions: [
          //           MeArchitectureOptions.newApi(),
          //           MeArchitectureOptions.apiSpec(),
          //           MeArchitectureOptions.botPlugin(),
          //         ],
          //         default: MeArchitectureOptions.newApi().id,
          //         placeholder: getLocalizedString(
          //           "core.createProjectQuestion.projectType.copilotExtension.placeholder"
          //         ),
          //         forgetLastValue: true,
          //         skipSingleOption: true,
          //         onDidSelection: onDidSelectionCapability,
          //       },
          //     },
          //   ],
          // },
          // {
          //   // 2.6 Office Add-in
          //   condition: {
          //     enum: [
          //       ProjectTypeOptions.officeMetaOS().id,
          //       ProjectTypeOptions.officeAddin().id,
          //       ProjectTypeOptions.outlookAddin().id,
          //     ],
          //   },
          //   data: {
          //     name: QuestionNames.Capabilities,
          //     title: (inputs: Inputs) => {
          //       const projectType = inputs[QuestionNames.ProjectType];
          //       switch (projectType) {
          //         case ProjectTypeOptions.outlookAddin().id:
          //           return getLocalizedString(
          //             "core.createProjectQuestion.projectType.outlookAddin.title"
          //           );
          //         case ProjectTypeOptions.officeMetaOS().id:
          //         case ProjectTypeOptions.officeAddin().id:
          //           return getLocalizedString(
          //             "core.createProjectQuestion.projectType.officeAddin.title"
          //           );
          //         default:
          //           return getLocalizedString("core.createCapabilityQuestion.titleNew");
          //       }
          //     },
          //     type: "singleSelect",
          //     staticOptions: [
          //       CapabilityOptions.jsonTaskPane(),
          //       ...(featureFlagManager.getBooleanValue(FeatureFlags.OfficeMetaOS)
          //         ? [CapabilityOptions.officeAddinImport()]
          //         : featureFlagManager.getBooleanValue(FeatureFlags.OfficeAddin)
          //         ? [CapabilityOptions.officeContentAddin(), CapabilityOptions.officeAddinImport()]
          //         : [CapabilityOptions.outlookAddinImport()]),
          //     ],
          //     placeholder: getLocalizedString("core.createCapabilityQuestion.placeholder"),
          //     forgetLastValue: true,
          //   },
          //   children: [
          //     {
          //       // office addin import sub-tree (capabilities=office-addin-import | outlook-addin-import)
          //       // TODO
          //       condition: {
          //         enum: [
          //           CapabilityOptions.outlookAddinImport().id,
          //           CapabilityOptions.officeAddinImport().id,
          //         ],
          //       },
          //       data: { type: "group", name: QuestionNames.OfficeAddinImport },
          //       children: [
          //         {
          //           data: {
          //             type: "folder",
          //             name: QuestionNames.OfficeAddinFolder,
          //             title: "Existing add-in project folder",
          //           },
          //         },
          //         {
          //           data: {
          //             type: "singleFile",
          //             name: QuestionNames.OfficeAddinManifest,
          //             title: "Select import project manifest file",
          //           },
          //         },
          //       ],
          //     },
          //     {
          //       // Office addin framework for json manifest
          //       condition: (inputs: Inputs) => {
          //         return (
          //           inputs[QuestionNames.ProjectType] === ProjectTypeOptions.officeAddin().id &&
          //           inputs[QuestionNames.Capabilities] !== CapabilityOptions.officeAddinImport().id
          //         );
          //       },
          //       data: {
          //         type: "singleSelect",
          //         name: QuestionNames.OfficeAddinFramework,
          //         title: getLocalizedString(
          //           "core.createProjectQuestion.projectType.officeAddin.framework.title"
          //         ),
          //         dynamicOptions: (inputs: Inputs) => {
          //           const projectType = inputs[QuestionNames.ProjectType];
          //           const capabilities = inputs[QuestionNames.Capabilities];
          //           if (projectType === ProjectTypeOptions.outlookAddin().id) {
          //             return [{ id: "default", label: "Default" }];
          //           } else if (
          //             (projectType === ProjectTypeOptions.officeAddin().id &&
          //               capabilities === CapabilityOptions.officeContentAddin().id) ||
          //             capabilities === CapabilityOptions.officeAddinImport().id
          //           ) {
          //             return [{ id: "default", label: "Default" }];
          //           } else {
          //             return [
          //               { id: "default", label: "Default" },
          //               { id: "react", label: "React" },
          //             ];
          //           }
          //         },
          //         staticOptions: [
          //           { id: "default", label: "Default" },
          //           { id: "react", label: "React" },
          //         ],
          //         placeholder: getLocalizedString(
          //           "core.createProjectQuestion.projectType.officeAddin.framework.placeholder"
          //         ),
          //         skipSingleOption: true,
          //       },
          //     },
          //   ],
          // },
          // {
          //   condition: (inputs: Inputs) => {
          //     return (
          //       inputs[QuestionNames.MeArchitectureType] == MeArchitectureOptions.newApi().id ||
          //       inputs[QuestionNames.ApiPluginType] == ApiPluginStartOptions.newApi().id
          //     );
          //   },
          //   data: {
          //     type: "singleSelect",
          //     name: QuestionNames.ApiAuth,
          //     title: getLocalizedString("core.createProjectQuestion.apiMessageExtensionAuth.title"),
          //     placeholder: getLocalizedString(
          //       "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
          //     ),
          //     staticOptions: [],
          //     dynamicOptions: (inputs: Inputs) => {
          //       const options: OptionItem[] = [ApiAuthOptions.none()];
          //       if (
          //         inputs[QuestionNames.MeArchitectureType] === MeArchitectureOptions.newApi().id
          //       ) {
          //         options.push(ApiAuthOptions.apiKey(), ApiAuthOptions.microsoftEntra());
          //       } else if (
          //         inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.newApi().id
          //       ) {
          //         options.push(ApiAuthOptions.apiKey());
          //         if (featureFlagManager.getBooleanValue(FeatureFlags.ApiPluginAAD)) {
          //           options.push(ApiAuthOptions.microsoftEntra());
          //         }
          //         options.push(ApiAuthOptions.oauth());
          //       }
          //       return options;
          //     },
          //     default: ApiAuthOptions.none().id,
          //   },
          // },
          // {
          //   // from API spec
          //   condition: (inputs: Inputs) => {
          //     return (
          //       (inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id ||
          //         inputs[QuestionNames.MeArchitectureType] === MeArchitectureOptions.apiSpec().id ||
          //         inputs[QuestionNames.CustomCopilotRag] ===
          //           CustomCopilotRagOptions.customApi().id) &&
          //       !(
          //         // Only skip this project when need to rediect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec
          //         (
          //           featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
          //           inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
          //           inputs[QuestionNames.Capabilities] === CapabilityOptions.declarativeAgent().id
          //         )
          //       )
          //     );
          //   },
          //   data: { type: "group", name: QuestionNames.FromExistingApi },
          //   children: [
          //     {
          //       data: apiSpecLocationQuestion(),
          //     },
          //     {
          //       data: apiOperationQuestion(),
          //       condition: (inputs: Inputs) => {
          //         return !inputs[QuestionNames.ApiPluginManifestPath];
          //       },
          //     },
          //   ],
          // },
        ],
      },
      {
        // language
        condition: (inputs: Inputs) => {
          const templateName = inputs[QuestionNames.TemplateName];
          const languages = Templates.filter((t) => t.name === templateName)
            .map((t) => t.language)
            .filter((lang) => lang !== "none" && lang !== undefined);
          return languages.length > 0;
        },
        data: {
          type: "singleSelect",
          title: getLocalizedString("core.ProgrammingLanguageQuestion.title"),
          name: QuestionNames.ProgrammingLanguage,
          staticOptions: [
            { id: ProgrammingLanguage.JS, label: "JavaScript" },
            { id: ProgrammingLanguage.TS, label: "TypeScript" },
            { id: ProgrammingLanguage.CSharp, label: "C#" },
            { id: ProgrammingLanguage.PY, label: "Python" },
          ],
          dynamicOptions: (inputs: Inputs) => {
            const templateName = inputs[QuestionNames.TemplateName];
            const languages = Templates.filter((t) => t.name === templateName)
              .map((t) => t.language)
              .filter((lang) => lang !== "none" && lang !== undefined);
            return languages;
          },
          skipSingleOption: true,
        },
      },
      {
        condition: (inputs: Inputs) => {
          // Only skip this project when need to rediect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec 3. No plugin manifest path
          return !(
            featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
            inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
            inputs[QuestionNames.Capabilities] === CapabilityOptions.declarativeAgent().id &&
            !inputs[QuestionNames.ApiPluginManifestPath]
          );
        },
        data: {
          type: "group",
        },
        children: [
          {
            //root folder
            data: {
              type: "folder",
              name: QuestionNames.Folder,
              title: getLocalizedString("core.question.workspaceFolder.title"),
              placeholder: getLocalizedString("core.question.workspaceFolder.placeholder"),
              default: path.join(os.homedir(), ConstantString.RootFolder),
            },
          },
          {
            //app name
            data: {
              type: "text",
              name: QuestionNames.AppName,
              title: getLocalizedString("core.question.appName.title"),
              default: async (inputs: Inputs) => {
                let defaultName = undefined;
                if (inputs[QuestionNames.SPFxSolution] == "import") {
                  defaultName = await getSolutionName(inputs[QuestionNames.SPFxFolder]);
                }
                return defaultName;
              },
              validation: {
                validFunc: async (
                  input: string,
                  previousInputs?: Inputs
                ): Promise<string | undefined> => {
                  const schema = {
                    pattern: AppNamePattern,
                    maxLength: 30,
                  };
                  if (input.length === 25) {
                    // show warning notification because it may exceed the Teams app name max length after appending suffix
                    const context = createContext();
                    void context.userInteraction.showMessage(
                      "warn",
                      getLocalizedString("core.QuestionAppName.validation.lengthWarning"),
                      false
                    );
                  }
                  const appName = input;
                  const validateResult = jsonschema.validate(appName, schema);
                  if (validateResult.errors && validateResult.errors.length > 0) {
                    if (validateResult.errors[0].name === "pattern") {
                      return getLocalizedString("core.QuestionAppName.validation.pattern");
                    }
                    if (validateResult.errors[0].name === "maxLength") {
                      return getLocalizedString("core.QuestionAppName.validation.maxlength");
                    }
                  }
                  if (previousInputs && previousInputs.folder) {
                    const folder = previousInputs.folder as string;
                    if (folder) {
                      const projectPath = path.resolve(folder, appName);
                      const exists = await fs.pathExists(projectPath);
                      if (exists)
                        return getLocalizedString(
                          "core.QuestionAppName.validation.pathExist",
                          projectPath
                        );
                    }
                  }
                  return undefined;
                },
              },
              placeholder: getLocalizedString("core.question.appName.placeholder"),
            },
          },
        ],
      },
    ],
  };
  return node;
}

export class ProjectTypeOptions {
  static tabOptionId = "tab-type";
  static botOptionId = "bot-type";
  static meOptionId = "me-type";
  static outlookAddinOptionId = "outlook-addin-type";
  static officeMetaOSOptionId = "office-meta-os-type";
  static officeAddinOptionId = "office-addin-type";
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
    return {
      id: ProjectTypeOptions.officeAddinOptionId,
      label: `$(extensions) ${getLocalizedString(
        "core.createProjectQuestion.projectType.officeAddin.label"
      )}`,
      detail: getLocalizedString("core.createProjectQuestion.projectType.officeAddin.detail"),
      groupName: ProjectTypeOptions.createGroupName,
    };
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

export class BotCapabilityOptions {
  static readonly basicBotId = "bot";
  static readonly notificationBotId = "notification";
  static readonly commandBotId = "command-bot";
  static readonly workflowBotId = "workflow-bot";

  static basicBot(): OptionItem {
    return {
      id: BotCapabilityOptions.basicBotId,
      label: `${getLocalizedString("core.BotNewUIOption.label")}`,
      detail: getLocalizedString("core.BotNewUIOption.detail"),
      data: TemplateNames.DefaultBot,
    };
  }

  // need further sub-options to decide template name
  static notificationBot(): OptionItem {
    return {
      id: BotCapabilityOptions.notificationBotId,
      label: `${getLocalizedString("core.NotificationOption.label")}`,
      detail: getLocalizedString("core.NotificationOption.detail"),
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }
  static commandBot(): OptionItem {
    return {
      // id must match cli `yargsHelp`
      id: BotCapabilityOptions.commandBotId,
      label: `${getLocalizedString("core.CommandAndResponseOption.label")}`,
      detail: getLocalizedString("core.CommandAndResponseOption.detail"),
      data: TemplateNames.CommandAndResponse,
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  static workflowBot(): OptionItem {
    const item: OptionItem = {
      id: BotCapabilityOptions.workflowBotId,
      label: `${getLocalizedString("core.WorkflowOption.label")}`,
      detail: getLocalizedString("core.WorkflowOption.detail"),
      data: TemplateNames.Workflow,
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
    return item;
  }
}

export class CapabilityOptions {
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

  static m365SsoLaunchPage(): OptionItem {
    return {
      id: "sso-launch-page",
      label: `${getLocalizedString("core.M365SsoLaunchPageOptionItem.label")}`,
      detail: getLocalizedString("core.M365SsoLaunchPageOptionItem.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.SsoTabObo,
    };
  }

  static dashboardTab(): OptionItem {
    return {
      id: "dashboard-tab",
      label: "core.DashboardOption.label",
      detail: getLocalizedString("core.DashboardOption.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.DashboardTab,
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  // need further sub-options to decide template name
  static SPFxTab(): OptionItem {
    return {
      id: "tab-spfx",
      label: getLocalizedString("core.TabSPFxOption.labelNew"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      detail: getLocalizedString("core.TabSPFxOption.detailNew"),
    };
  }

  // need further sub-options to decide template name
  static m365SearchMe(): OptionItem {
    return {
      id: "search-app",
      label: `${getLocalizedString("core.M365SearchAppOptionItem.label")}`,
      detail: getLocalizedString("core.M365SearchAppOptionItem.copilot.detail"),
    };
  }

  static collectFormMe(): OptionItem {
    return {
      id: "collect-form-message-extension",
      label: `${getLocalizedString("core.MessageExtensionOption.labelNew")}`,
      detail: getLocalizedString("core.MessageExtensionOption.detail"),
      data: TemplateNames.MessageExtensionAction,
    };
  }

  static linkUnfurling(): OptionItem {
    return {
      id: "link-unfurling",
      label: `${getLocalizedString("core.LinkUnfurlingOption.label")}`,
      detail: getLocalizedString("core.LinkUnfurlingOption.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlook"
      ),
      data: TemplateNames.LinkUnfurling,
    };
  }

  static outlookAddinImport(): OptionItem {
    return {
      id: "outlook-addin-import",
      label: getLocalizedString("core.importAddin.label"),
      detail: getLocalizedString("core.importAddin.detail"),
    };
  }
  static officeContentAddin(): OptionItem {
    return {
      id: "office-content-addin",
      label: getLocalizedString("core.officeContentAddin.label"),
      detail: getLocalizedString("core.officeContentAddin.detail"),
    };
  }
  static officeAddinImport(): OptionItem {
    return {
      id: "office-addin-import",
      label: getLocalizedString("core.importOfficeAddin.label"),
      detail: getLocalizedString("core.importAddin.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.previewOnWindow"
      ),
    };
  }
  static jsonTaskPane(): OptionItem {
    return {
      id: "json-taskpane",
      label: getLocalizedString("core.newTaskpaneAddin.label"),
      detail: getLocalizedString("core.newTaskpaneAddin.detail"),
    };
  }

  static apiPlugin(): OptionItem {
    return {
      id: "api-plugin",
      label: getLocalizedString("core.createProjectQuestion.projectType.copilotPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.copilotPlugin.detail"),
    };
  }

  static declarativeAgent(): OptionItem {
    return {
      id: "declarative-agent",
      label: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.detail"),
    };
  }

  // custom copilot
  static customCopilotBasic(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
      ? getLocalizedString("core.createProjectQuestion.capability.customEngineAgent.description")
      : undefined;
    return {
      id: "custom-copilot-basic",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotBasicOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotBasicOption.detail"
      ),
      description: description,
    };
  }

  static customCopilotRag(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
      ? getLocalizedString("core.createProjectQuestion.capability.customEngineAgent.description")
      : undefined;
    return {
      id: "custom-copilot-rag",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.detail"
      ),
      description: description,
    };
  }

  static customCopilotAssistant(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
      ? getLocalizedString("core.createProjectQuestion.capability.customEngineAgent.description")
      : undefined;
    return {
      id: "custom-copilot-agent",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantOption.detail"
      ),
      description: description,
    };
  }
}

export class NotificationBotOptions {
  static appService(): HostTypeTriggerOptionItem {
    return {
      id: "http-express",
      hostType: HostType.AppService,
      label: getLocalizedString("plugins.bot.triggers.http-express.label"),
      description: getLocalizedString("plugins.bot.triggers.http-express.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-express.detail"),
      data: TemplateNames.NotificationExpress,
    };
  }

  static functionsTimerTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "timer-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.timer-functions.detail"),
      data: TemplateNames.NotificationTimerTrigger,
    };
  }

  static functionsHttpAndTimerTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "http-and-timer-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP, NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.detail"),
      data: TemplateNames.NotificationHttpTimerTrigger,
    };
  }

  static functionsHttpTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "http-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP],
      label: getLocalizedString("plugins.bot.triggers.http-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-functions.detail"),
      data: TemplateNames.NotificationHttpTrigger,
    };
  }
}

export class MeArchitectureOptions {
  static botPlugin(): OptionItem {
    return {
      id: "bot-plugin",
      label: getLocalizedString("core.createProjectQuestion.capability.botMessageExtension.label"),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.botMessageExtension.detail"
      ),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookCopilot"
      ),
      data: TemplateNames.MessageExtensionCopilot,
    };
  }

  static newApi(): OptionItem {
    return {
      id: "new-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.messageExtensionNewApiOption.detail"
      ),
    };
  }

  static apiSpec(): OptionItem {
    return {
      id: "api-spec",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.messageExtensionApiSpecOption.detail"
      ),
    };
  }
}

export class DeclarativeCopilotTypeOptions {
  static noPlugin(): OptionItem {
    return {
      id: "no",
      label: getLocalizedString("core.createProjectQuestion.noPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.noPlugin.detail"),
      data: TemplateNames.BasicGpt,
    };
  }
  static withPlugin(): OptionItem {
    return {
      id: "yes",
      label: getLocalizedString("core.createProjectQuestion.addPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.addPlugin.detail"),
    };
  }
}

export class ApiPluginStartOptions {
  static newApi(): OptionItem {
    return {
      id: "new-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.detail"
      ),
    };
  }

  static apiSpec(): OptionItem {
    return {
      id: "api-spec",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.detail"
      ),
      data: TemplateNames.DeclarativeAgentWithApiSpec,
    };
  }

  static existingPlugin(): OptionItem {
    return {
      id: "existing-plugin",
      label: getLocalizedString("core.createProjectQuestion.apiPlugin.importPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.apiPlugin.importPlugin.detail"),
    };
  }
}

export class ApiAuthOptions {
  static none(): OptionItem {
    return {
      id: "none",
      label: "None",
      data: TemplateNames.ApiPluginFromScratch,
    };
  }
  static apiKey(): OptionItem {
    return {
      id: "api-key",
      label: "API Key (Bearer Token Auth)",
      data: TemplateNames.ApiPluginFromScratchBearer,
    };
  }

  static microsoftEntra(): OptionItem {
    return {
      id: "microsoft-entra",
      label: "Microsoft Entra",
    };
  }

  static oauth(): OptionItem {
    return {
      id: "oauth",
      label: "OAuth",
      data: TemplateNames.ApiPluginFromScratchOAuth,
    };
  }

  static all(): OptionItem[] {
    return [
      ApiAuthOptions.none(),
      ApiAuthOptions.apiKey(),
      ApiAuthOptions.microsoftEntra(),
      ApiAuthOptions.oauth(),
    ];
  }
}

export class CustomCopilotRagOptions {
  static customize(): OptionItem {
    return {
      id: "custom-copilot-rag-customize",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.detail"
      ),
    };
  }

  static azureAISearch(): OptionItem {
    return {
      id: "custom-copilot-rag-azureAISearch",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.detail"
      ),
    };
  }

  static customApi(): OptionItem {
    return {
      id: "custom-copilot-rag-customApi",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.detail"
      ),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
    };
  }

  static microsoft365(): OptionItem {
    return {
      id: "custom-copilot-rag-microsoft365",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagMicrosoft365Option.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagMicrosoft365Option.detail"
      ),
    };
  }
}

export class CustomCopilotAssistantOptions {
  static new(): OptionItem {
    return {
      id: "custom-copilot-agent-new",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantNewOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantNewOption.detail"
      ),
    };
  }

  static assistantsApi(): OptionItem {
    return {
      id: "custom-copilot-agent-assistants-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantAssistantsApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantAssistantsApiOption.detail"
      ),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
    };
  }
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

export async function getSolutionName(spfxFolder: string): Promise<string | undefined> {
  const yoInfoPath = path.join(spfxFolder, Constants.YO_RC_FILE);
  if (await fs.pathExists(yoInfoPath)) {
    const yoInfo = await fs.readJson(yoInfoPath);
    if (yoInfo["@microsoft/generator-sharepoint"]) {
      return yoInfo["@microsoft/generator-sharepoint"][Constants.YO_RC_SOLUTION_NAME];
    } else {
      return undefined;
    }
  } else {
    throw new FileNotFoundError(Constants.PLUGIN_NAME, yoInfoPath, Constants.IMPORT_HELP_LINK);
  }
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
