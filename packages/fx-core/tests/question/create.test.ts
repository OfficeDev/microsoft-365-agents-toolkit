// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ErrorType, SpecParser, ValidationStatus, WarningType } from "@microsoft/m365-spec-parser";
import {
  Context,
  FuncValidation,
  FxError,
  Inputs,
  LocalFunc,
  LogProvider,
  MultiSelectQuestion,
  OptionItem,
  Platform,
  Question,
  SingleFileQuestion,
  SingleSelectQuestion,
  UserError,
  UserInteraction,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { assert, expect } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as path from "path";
import sinon from "sinon";
import { FeatureFlagName } from "../../src/common/featureFlags";
import * as utils from "../../src/common/globalVars";
import { setTools } from "../../src/common/globalVars";
import { getLocalizedString } from "../../src/common/localizeUtils";
import { sampleProvider } from "../../src/common/samples";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { convertToLangKey } from "../../src/component/generator/utils";
import { FileNotFoundError } from "../../src/error";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  CapabilityOptions,
  CustomCopilotAssistantOptions,
  CustomCopilotRagOptions,
  DeclarativeCopilotTypeOptions,
  MeArchitectureOptions,
  NotificationTriggerOptions,
  ProgrammingLanguage,
  ProjectTypeOptions,
  QuestionNames,
  RuntimeOptions,
  SPFxVersionOptionIds,
  apiAuthQuestion,
  apiOperationQuestion,
  apiSpecLocationQuestion,
  appNameQuestion,
  capabilityQuestion,
  createProjectQuestionNode,
  createSampleProjectQuestionNode,
  folderQuestion,
  getLanguageOptions,
  getSolutionName,
  pluginApiSpecQuestion,
  pluginManifestQuestion,
  programmingLanguageQuestion,
} from "../../src/question";
import { QuestionTreeVisitor, traverse } from "../../src/ui/visitor";
import { MockTools, MockUserInteraction, randomAppName } from "../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../plugins/solution/util";

export async function callFuncs(question: Question, inputs: Inputs, answer?: string) {
  try {
    if (question.default && typeof question.default !== "string") {
      await (question.default as LocalFunc<string | undefined>)(inputs);
    }

    if (
      (question.type === "singleSelect" || question.type === "multiSelect") &&
      typeof question.dynamicOptions !== "object" &&
      question.dynamicOptions
    ) {
      await question.dynamicOptions(inputs);
    }
    if (answer && (question as any).validation?.validFunc) {
      await (question as any).validation.validFunc(answer, inputs);
    }

    if ((question as any).placeholder && typeof (question as any).placeholder !== "string") {
      await (question as any).placeholder(inputs);
    }
  } catch (e) {}
}

describe("scaffold question", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("createProjectQuestionNode", () => {
    const ui = new MockUserInteraction();
    let mockedEnvRestore: RestoreFn = () => {};

    beforeEach(() => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.SampleConfigBranch]: "dev",
        [FeatureFlagName.ChatParticipantUIEntries]: "false",
      });
    });
    afterEach(() => {
      mockedEnvRestore();
    });

    it("create sample", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.Samples) {
          return ok({ type: "success", result: "abc" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      sandbox.stub(sampleProvider, "SampleCollection").resolves({ samples: ["1"] });
      await traverse(createSampleProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [QuestionNames.Samples, QuestionNames.Folder]);
    });

    it("traverse in vscode notification bot", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 6);
          assert.isNotEmpty((options as OptionItem[])[0].groupName);
          return ok({ type: "success", result: ProjectTypeOptions.bot().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 4);
          const title =
            typeof question.title === "function" ? await question.title(inputs) : question.title;
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.projectType.bot.title")
          );
          return ok({ type: "success", result: CapabilityOptions.notificationBot().id });
        } else if (question.name === QuestionNames.BotTrigger) {
          return ok({ type: "success", result: NotificationTriggerOptions.appService().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.BotTrigger,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode bot me", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 6);
          assert.isFalse((options[2] as OptionItem).detail?.includes("Copilot"));
          return ok({ type: "success", result: ProjectTypeOptions.me().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          const title =
            typeof question.title === "function" ? await question.title(inputs) : question.title;
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.projectType.messageExtension.title")
          );
          return ok({ type: "success", result: CapabilityOptions.m365SearchMe().id });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const options = await (question as SingleSelectQuestion).dynamicOptions!(inputs);
          assert.deepEqual(options, MeArchitectureOptions.all());
          return ok({ type: "success", result: MeArchitectureOptions.botMe().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.MeArchitectureType,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode me from new api (No authentication)", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.me().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          const title =
            typeof question.title === "function" ? await question.title(inputs) : question.title;
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.projectType.messageExtension.title")
          );
          return ok({ type: "success", result: CapabilityOptions.m365SearchMe().id });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: MeArchitectureOptions.newApi().id });
        } else if (question.name === QuestionNames.ApiAuth) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions?.(inputs);
          assert.isTrue(options?.length === 3);
          return ok({ type: "success", result: ApiAuthOptions.none().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.staticOptions;
          // Assert
          assert.equal(options.length, 2);
          // Assert
          assert.equal(options.length, 2);
          assert.deepEqual(options, [
            MeArchitectureOptions.newApi(),
            MeArchitectureOptions.apiSpec(),
          ]);
          return ok({ type: "success", result: MeArchitectureOptions.newApi().id });
        } else if (question.name === QuestionNames.ApiAuth) {
          const select = question as SingleSelectQuestion;
          const options = select.staticOptions;
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: ApiAuthOptions.none().id });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.MeArchitectureType,
        QuestionNames.ApiAuth,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode me from new api (key auth)", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.me().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          const title =
            typeof question.title === "function" ? await question.title(inputs) : question.title;
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.projectType.messageExtension.title")
          );
          return ok({ type: "success", result: CapabilityOptions.m365SearchMe().id });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: MeArchitectureOptions.newApi().id });
        } else if (question.name === QuestionNames.ApiAuth) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions?.(inputs);
          assert.isTrue(options?.length === 3);
          return ok({ type: "success", result: ApiAuthOptions.apiKey().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.MeArchitectureType,
        QuestionNames.ApiAuth,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode me from new api (sso auth)", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.me().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          const title =
            typeof question.title === "function" ? await question.title(inputs) : question.title;
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.projectType.messageExtension.title")
          );
          return ok({ type: "success", result: CapabilityOptions.m365SearchMe().id });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: MeArchitectureOptions.newApi().id });
        } else if (question.name === QuestionNames.ApiAuth) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions?.(inputs);
          assert.isTrue(options?.length === 3);
          return ok({
            type: "success",
            result: ApiAuthOptions.microsoftEntra().id,
          });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.MeArchitectureType,
        QuestionNames.ApiAuth,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode api me from existing api", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.me().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: CapabilityOptions.m365SearchMe().id });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: MeArchitectureOptions.apiSpec().id });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.staticOptions;
          // Assert
          assert.equal(options.length, 2);
          return ok({ type: "success", result: MeArchitectureOptions.apiSpec().id });
        } else if (question.name === QuestionNames.ApiSpecLocation) {
          inputs.supportedApisFromApiSpec = [
            { id: "operation1", label: "operation1", groupName: "1" },
            { id: "operation2", label: "operation2", groupName: "2" },
          ];
          return ok({ type: "success", result: "https://test.com" });
        } else if (question.name === QuestionNames.ApiOperation) {
          return ok({ type: "success", result: ["operation1"] });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.MeArchitectureType,
        QuestionNames.ApiSpecLocation,
        QuestionNames.ApiOperation,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode Outlook addin import", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          return ok({ type: "success", result: ProjectTypeOptions.outlookAddin().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = (await select.dynamicOptions!(inputs)) as OptionItem[];
          assert.deepEqual(
            options.map((o) => o.id),
            [CapabilityOptions.officeAddinTaskpane().id, CapabilityOptions.outlookAddinImport().id]
          );
          return ok({ type: "success", result: CapabilityOptions.outlookAddinImport().id });
        } else if (question.name === QuestionNames.OfficeAddinFolder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.OfficeAddinManifest) {
          return ok({ type: "success", result: "./manifest.json" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.OfficeAddinFolder,
        QuestionNames.OfficeAddinManifest,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in vscode Office meta os", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.officeMetaOS().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          const items = CapabilityOptions.officeAddinCapabilities(
            ProjectTypeOptions.officeMetaOS().id
          );
          assert.deepEqual(options, items);
          return ok({ type: "success", result: CapabilityOptions.officeAddinImport().id });
        } else if (question.name === QuestionNames.OfficeAddinFolder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.OfficeAddinManifest) {
          return ok({ type: "success", result: "./manifest.json" });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: "typescript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.OfficeAddinFramework) {
          return ok({ type: "success", result: "default" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.OfficeAddinFolder,
        QuestionNames.OfficeAddinManifest,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });
    it("traverse in vscode Office meta os v2", async () => {
      const innerMockedEnvRestore = mockedEnv({
        [FeatureFlagName.OfficeMetaOS]: "true",
      });
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.officeMetaOS().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          const items = CapabilityOptions.officeAddinCapabilities(
            ProjectTypeOptions.officeMetaOS().id
          );
          assert.deepEqual(options, items);
          return ok({ type: "success", result: "json-taskpane" });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
      if (innerMockedEnvRestore) {
        innerMockedEnvRestore();
      }
    });
    it("traverse in vscode SPFx new", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.tab().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 4);
          const title =
            typeof question.title === "function" ? await question.title(inputs) : question.title;
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.projectType.tab.title")
          );
          return ok({ type: "success", result: CapabilityOptions.SPFxTab().id });
        } else if (question.name === QuestionNames.SPFxSolution) {
          return ok({ type: "success", result: "new" });
        } else if (question.name === QuestionNames.SPFxInstallPackage) {
          return ok({ type: "success", result: SPFxVersionOptionIds.installLocally });
        } else if (question.name === QuestionNames.SPFxFramework) {
          return ok({ type: "success", result: "react" });
        } else if (question.name === QuestionNames.SPFxWebpartName) {
          return ok({ type: "success", result: "test" });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: "typescript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.SPFxSolution,
        QuestionNames.SPFxInstallPackage,
        QuestionNames.SPFxFramework,
        QuestionNames.SPFxWebpartName,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });
    it("traverse in vscode SPFx import", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.tab().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 4);
          return ok({ type: "success", result: CapabilityOptions.SPFxTab().id });
        } else if (question.name === QuestionNames.SPFxSolution) {
          return ok({ type: "success", result: "import" });
        } else if (question.name === QuestionNames.SPFxFolder) {
          return ok({ type: "success", result: "" });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: "typescript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.SPFxSolution,
        QuestionNames.SPFxFolder,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });
    it("traverse in vscode TDP with tab and bot", async () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        staticTabs: [
          {
            name: "tab1",
            entityId: "tab1",
            contentUrl: "mock-contentUrl",
            websiteUrl: "mock-websiteUrl",
            context: [],
            scopes: [],
          },
        ],
        bots: [
          {
            botId: "mock-bot-id",
            isNotificationOnly: false,
            needsChannelSelector: false,
            supportsCalling: false,
            supportsFiles: false,
            supportsVideo: false,
            scopes: [],
            teamCommands: [],
            groupChatCommands: [],
            personalCommands: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        teamsAppFromTdp: appDefinition,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: "tab-bot-type" });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: CapabilityOptions.nonSsoTabAndBot().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 2);
          return ok({ type: "success", result: "typescript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.ReplaceWebsiteUrl) {
          const select = question as MultiSelectQuestion;
          const options = (await select.dynamicOptions!(inputs)) as OptionItem[];
          const defaults = options.map((o) => o.id);
          assert.isTrue(options.length === 1);
          assert.isTrue(defaults.length === 1);
          return ok({ type: "success", result: [] });
        } else if (question.name === QuestionNames.ReplaceContentUrl) {
          const select = question as MultiSelectQuestion;
          const options = (await select.dynamicOptions!(inputs)) as OptionItem[];
          const defaults = options.map((o) => o.id);
          assert.isTrue(options.length === 1);
          assert.isTrue(defaults.length === 1);
          return ok({ type: "success", result: [] });
        } else if (question.name === QuestionNames.ReplaceBotIds) {
          const select = question as MultiSelectQuestion;
          const options = (await select.dynamicOptions!(inputs)) as OptionItem[];
          const defaults = options.map((o) => o.id);
          assert.isTrue(options.length === 1);
          assert.isTrue(defaults.length === 1);
          return ok({ type: "success", result: [] });
        }
        return ok({ type: "success", result: undefined });
      };
      const tres = await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.isTrue(tres.isOk());
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
        QuestionNames.ReplaceWebsiteUrl,
        QuestionNames.ReplaceContentUrl,
        QuestionNames.ReplaceBotIds,
      ]);
    });
    it("traverse in vscode TDP with empty website url", async () => {
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        staticTabs: [
          {
            name: "tab1",
            entityId: "tab1",
            contentUrl: "https://test.com",
            websiteUrl: "",
            context: [],
            scopes: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        teamsAppFromTdp: appDefinition,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: "tab-bot-type" });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: CapabilityOptions.nonSsoTabAndBot().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 2);
          return ok({ type: "success", result: "typescript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.ReplaceContentUrl) {
          const select = question as MultiSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 1);
          return ok({ type: "success", result: [] });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
        QuestionNames.ReplaceContentUrl,
      ]);
    });
    it("traverse in cli", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "false" });
      const inputs: Inputs = {
        platform: Platform.CLI,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.Capabilities) {
          return ok({ type: "success", result: CapabilityOptions.notificationBot().id });
        } else if (question.name === QuestionNames.BotTrigger) {
          return ok({ type: "success", result: NotificationTriggerOptions.appService().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.BotTrigger,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("traverse in cli TEAMSFX_CLI_DOTNET=true", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "true" });
      const inputs: Inputs = {
        platform: Platform.CLI,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.Runtime) {
          return ok({ type: "success", result: RuntimeOptions.DotNet().id });
        } else if (question.name === QuestionNames.Capabilities) {
          return ok({ type: "success", result: CapabilityOptions.notificationBot().id });
        } else if (question.name === QuestionNames.BotTrigger) {
          return ok({ type: "success", result: NotificationTriggerOptions.appServiceForVS().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.Runtime,
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.BotTrigger,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    describe("Custom Copilot", () => {
      let mockedEnvRestore: RestoreFn;
      const tools = new MockTools();
      setTools(tools);

      afterEach(() => {
        if (mockedEnvRestore) {
          mockedEnvRestore();
        }
      });

      it("Basic AI Chatbot - OpenAI", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotBasic().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-openai" });
          } else if (question.name === QuestionNames.OpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.OpenAIKey,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("RAG - Customize - Azure OpenAI", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotRag().id });
          } else if (question.name === QuestionNames.CustomCopilotRag) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: CustomCopilotRagOptions.customize().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "python" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-azure-openai" });
          } else if (question.name === QuestionNames.AzureOpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.AzureOpenAIEndpoint) {
            return ok({ type: "success", result: "testEndppint" });
          } else if (question.name === QuestionNames.AzureOpenAIDeploymentName) {
            return ok({ type: "success", result: "testAzureOpenAIDeploymentName" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotRag,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.AzureOpenAIKey,
          QuestionNames.AzureOpenAIEndpoint,
          QuestionNames.AzureOpenAIDeploymentName,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("RAG - Customize - Azure OpenAI wit empty endpoint", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotRag().id });
          } else if (question.name === QuestionNames.CustomCopilotRag) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: CustomCopilotRagOptions.customize().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "python" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-azure-openai" });
          } else if (question.name === QuestionNames.AzureOpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotRag,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.AzureOpenAIKey,
          QuestionNames.AzureOpenAIEndpoint,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("RAG - Azure AI Search - Azure OpenAI", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotRag().id });
          } else if (question.name === QuestionNames.CustomCopilotRag) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: CustomCopilotRagOptions.customize().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "python" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-azure-openai" });
          } else if (question.name === QuestionNames.AzureOpenAIKey) {
            return ok({ type: "success", result: undefined });
          } else if (question.name === QuestionNames.AzureOpenAIEndpoint) {
            return ok({ type: "success", result: "testEndppint" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotRag,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.AzureOpenAIKey,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("RAG - Custom API - Azure OpenAI", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);

            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotRag().id });
          } else if (question.name === QuestionNames.CustomCopilotRag) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: CustomCopilotRagOptions.customApi().id });
          } else if (question.name === QuestionNames.ApiSpecLocation) {
            inputs.supportedApisFromApiSpec = [
              { id: "operation1", label: "operation1", groupName: "1" },
              { id: "operation2", label: "operation2", groupName: "2" },
            ];
            return ok({ type: "success", result: "https://test.com" });
          } else if (question.name === QuestionNames.ApiOperation) {
            return ok({ type: "success", result: ["operation1"] });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-azure-openai" });
          } else if (question.name === QuestionNames.AzureOpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.AzureOpenAIEndpoint) {
            return ok({ type: "success", result: "testEndppint" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotRag,
          QuestionNames.ApiSpecLocation,
          QuestionNames.ApiOperation,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.AzureOpenAIKey,
          QuestionNames.AzureOpenAIEndpoint,
          QuestionNames.AzureOpenAIDeploymentName,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("RAG - Microsoft 365 - Azure OpenAI", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);

            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotRag().id });
          } else if (question.name === QuestionNames.CustomCopilotRag) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: CustomCopilotRagOptions.microsoft365().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-azure-openai" });
          } else if (question.name === QuestionNames.AzureOpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.AzureOpenAIEndpoint) {
            return ok({ type: "success", result: "testEndppint" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotRag,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.AzureOpenAIKey,
          QuestionNames.AzureOpenAIEndpoint,
          QuestionNames.AzureOpenAIDeploymentName,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("AI Assistant - New - OpenAI", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);

            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotAssistant().id });
          } else if (question.name === QuestionNames.CustomCopilotAssistant) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: CustomCopilotAssistantOptions.new().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-openai" });
          } else if (question.name === QuestionNames.OpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotAssistant,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.OpenAIKey,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("AI Assistant - Assistants API", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);

            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotAssistant().id });
          } else if (question.name === QuestionNames.CustomCopilotAssistant) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({
              type: "success",
              result: CustomCopilotAssistantOptions.assistantsApi().id,
            });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-openai" });
          } else if (question.name === QuestionNames.OpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotAssistant,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.OpenAIKey,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("AI Assistant - Assistants API Python", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);

            return ok({ type: "success", result: ProjectTypeOptions.customCopilot().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: CapabilityOptions.customCopilotAssistant().id });
          } else if (question.name === QuestionNames.CustomCopilotAssistant) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({
              type: "success",
              result: CustomCopilotAssistantOptions.assistantsApi().id,
            });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: "python" });
          } else if (question.name === QuestionNames.LLMService) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "llm-service-azure-openai" });
          } else if (question.name === QuestionNames.AzureOpenAIKey) {
            return ok({ type: "success", result: "testKey" });
          } else if (question.name === QuestionNames.AzureOpenAIEndpoint) {
            return ok({ type: "success", result: "testEndppint" });
          } else if (question.name === QuestionNames.AzureOpenAIDeploymentName) {
            return ok({ type: "success", result: "testAzureOpenAIDeploymentName" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.CustomCopilotAssistant,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.LLMService,
          QuestionNames.AzureOpenAIKey,
          QuestionNames.AzureOpenAIEndpoint,
          QuestionNames.AzureOpenAIDeploymentName,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });
    });

    describe("copilot plugin enabled", () => {
      let mockedEnvRestore: RestoreFn;
      const tools = new MockTools();
      setTools(tools);
      beforeEach(() => {
        mockedEnvRestore = mockedEnv({});
      });

      afterEach(() => {
        if (mockedEnvRestore) {
          mockedEnvRestore();
        }
      });

      it("traverse in vscode Copilot Plugin from new API with auth enabled", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            return ok({ type: "success", result: CapabilityOptions.apiPlugin().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            const select = question as SingleSelectQuestion;
            const options = select.staticOptions;
            assert.isTrue(options.length === 3);
            return ok({ type: "success", result: ApiPluginStartOptions.newApi().id });
          } else if (question.name === QuestionNames.ApiAuth) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: ApiAuthOptions.none().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiPluginType,
          QuestionNames.ApiAuth,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("traverse in vscode Copilot Plugin from Kiota", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        inputs[QuestionNames.Capabilities] = CapabilityOptions.apiPlugin().id;
        inputs[QuestionNames.ApiSpecLocation] = "api-spec-path";
        inputs[QuestionNames.ApiPluginManifestPath] = "api-plugin-manifest-path";
        inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.apiSpec().id;
        inputs[QuestionNames.ApiOperation] = "api-plugin-manifest-path";
        inputs[QuestionNames.ProjectType] = ProjectTypeOptions.Agent().id;
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiSpecLocation,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("traverse in vscode Declarative Copilot from Kiota", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        inputs[QuestionNames.Capabilities] = CapabilityOptions.declarativeAgent().id;
        inputs[QuestionNames.ApiSpecLocation] = "api-spec-path";
        inputs[QuestionNames.ApiPluginManifestPath] = "api-plugin-manifest-path";
        inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.apiSpec().id;
        inputs[QuestionNames.ApiOperation] = "api-plugin-manifest-path";
        inputs[QuestionNames.ProjectType] = ProjectTypeOptions.Agent().id;
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiSpecLocation,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("traverse in vscode Copilot Plugin to Kiota", async () => {
        mockedEnvRestore = mockedEnv({
          [FeatureFlagName.KiotaIntegration]: "true",
        });
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };

        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            return ok({ type: "success", result: CapabilityOptions.apiPlugin().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            return ok({ type: "success", result: ApiPluginStartOptions.apiSpec().id });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiPluginType,
        ]);
      });

      it("traverse in vscode Declarative Copilot to Kiota", async () => {
        mockedEnvRestore = mockedEnv({
          [FeatureFlagName.KiotaIntegration]: "true",
        });
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };

        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            return ok({ type: "success", result: CapabilityOptions.declarativeAgent().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            return ok({ type: "success", result: ApiPluginStartOptions.apiSpec().id });
          } else if (question.name === QuestionNames.WithPlugin) {
            return ok({ type: "success", result: "yes" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.WithPlugin,
          QuestionNames.ApiPluginType,
        ]);
      });

      it("traverse in vscode Copilot Plugin from new API with API Key authentication", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            return ok({ type: "success", result: CapabilityOptions.apiPlugin().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            return ok({ type: "success", result: ApiPluginStartOptions.newApi().id });
          } else if (question.name === QuestionNames.ApiAuth) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 4);
            return ok({ type: "success", result: ApiAuthOptions.apiKey().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiPluginType,
          QuestionNames.ApiAuth,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("traverse in vscode Copilot Plugin from API Spec", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          if (question.name !== QuestionNames.ApiOperation) {
            await callFuncs(question, inputs);
          }
          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.equal(
              (question.title as any)!(inputs),
              getLocalizedString("core.createProjectQuestion.projectType.copilotExtension.title")
            );
            return ok({ type: "success", result: CapabilityOptions.apiPlugin().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            return ok({ type: "success", result: ApiPluginStartOptions.apiSpec().id });
          } else if (question.name === QuestionNames.ApiSpecLocation) {
            const validRes = await (question as any).inputBoxConfig.validation!.validFunc(
              "https://test.com"
            );
            assert.isUndefined(validRes);
            return ok({ type: "success", result: "https://test.com" });
          } else if (question.name === QuestionNames.ApiOperation) {
            return ok({ type: "success", result: ["testOperation1"] });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: "typescript" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiPluginType,
          QuestionNames.ApiSpecLocation,
          QuestionNames.ApiOperation,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("traverse in cli", async () => {
        mockedEnvRestore = mockedEnv({
          TEAMSFX_CLI_DOTNET: "false",
        });

        const inputs: Inputs = {
          platform: Platform.CLI,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);
          await callFuncs(question, inputs);
          if (question.name === QuestionNames.Capabilities) {
            return ok({ type: "success", result: CapabilityOptions.apiPlugin().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            return ok({ type: "success", result: ApiPluginStartOptions.newApi().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            return ok({ type: "success", result: "javascript" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          } else if (question.name === QuestionNames.ApiAuth) {
            return ok({ type: "success", result: ApiAuthOptions.none().id });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.ApiPluginType,
          QuestionNames.ApiAuth,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      describe("list operations", async () => {
        const mockedEnvRestore: RestoreFn = () => {};

        afterEach(() => {
          mockedEnvRestore();
        });
        it("list operations successfully", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const options = (await question.dynamicOptions!(inputs)) as OptionItem[];
          const placeholder = (question as any).placeholder(inputs) as string;
          const title = (question as any).title(inputs) as string;

          assert.isTrue(options.length === 2);
          assert.isTrue(options[0].id === "operation1");
          assert.isTrue(options[1].id === "operation2");
          assert.equal(
            placeholder,
            getLocalizedString("core.createProjectQuestion.apiSpec.operation.apikey.placeholder")
          );
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.apiSpec.operation.title")
          );
        });

        it("list operations for API plugin successfully", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
            [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const options = (await question.dynamicOptions!(inputs)) as OptionItem[];
          const placeholder = (question as any).placeholder(inputs) as string;
          const title = (question as any).title(inputs) as string;

          assert.isTrue(options.length === 2);
          assert.isTrue(options[0].id === "operation1");
          assert.isTrue(options[1].id === "operation2");
          assert.equal(
            placeholder,
            getLocalizedString("core.createProjectQuestion.apiSpec.operation.plugin.placeholder")
          );
          assert.equal(
            title,
            getLocalizedString("core.createProjectQuestion.apiSpec.copilotOperation.title")
          );
        });

        it("validate operations error: missing inputs", async () => {
          const question = apiOperationQuestion();

          const validationSchema = question.validation as FuncValidation<string[]>;

          let hasError = false;
          try {
            await validationSchema.validFunc!(["operation1", "operation2"], undefined);
          } catch (e) {
            hasError = true;
          }

          assert.isTrue(hasError);
        });

        it(" validate operations successfully", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  serverUrl: "https://server1",
                  authName: "oauth2",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                  authName: "oauth2",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(["operation1", "operation2"], inputs);

          assert.deepEqual(inputs.apiAuthData, [
            {
              serverUrl: "https://server1",
              authName: "oauth2",
            },
          ]);
          assert.isUndefined(res);
        });

        it(" validate operations successfully with Teams AI project", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            "custom-copilot-rag": "custom-copilot-rag-customApi",
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation3",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation4",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation5",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation6",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation7",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation8",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation9",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation10",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation11",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(
            [
              "operation1",
              "operation2",
              "operation3",
              "operation4",
              "operation5",
              "operation6",
              "operation7",
              "operation8",
              "operation9",
              "operation10",
              "operation11",
            ],
            inputs
          );

          assert.isUndefined(res);
        });

        it(" validate operations successfully with copilot project", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation3",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation4",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation5",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation6",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation7",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation8",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation9",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation10",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation11",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(
            [
              "operation1",
              "operation2",
              "operation3",
              "operation4",
              "operation5",
              "operation6",
              "operation7",
              "operation8",
              "operation9",
              "operation10",
              "operation11",
            ],
            inputs
          );

          assert.isUndefined(res);
        });

        it(" validate operations successfully due to length limitation for sme project", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            [QuestionNames.ProjectType]: ProjectTypeOptions.me().id,
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation3",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation4",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation5",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation6",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation7",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation8",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation9",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation10",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation11",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(
            [
              "operation1",
              "operation2",
              "operation3",
              "operation4",
              "operation5",
              "operation6",
              "operation7",
              "operation8",
              "operation9",
              "operation10",
              "operation11",
            ],
            inputs
          );

          expect(res).to.equal(
            "11 API(s) selected. You can select at least one and at most 10 APIs."
          );
        });

        it(" validate operations with auth successfully", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  authName: "auth1",
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  authName: "auth1",
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const placeholder = (question as any).placeholder(inputs) as string;
          const res = await validationSchema.validFunc!(["operation1", "operation2"], inputs);

          assert.isTrue(placeholder.includes("API key"));
          assert.isUndefined(res);
        });

        it(" validate operations should return error message when selected APIs with multiple server url", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  authName: "auth1",
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  authName: "auth1",
                  serverUrl: "https://server2",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(["operation1", "operation2"], inputs);

          assert.equal(
            res,
            getLocalizedString(
              "core.createProjectQuestion.apiSpec.operation.multipleServer",
              ["https://server1", "https://server2"].join(", ")
            )
          );
        });

        it(" validate operations should success when selected APIs with multiple server url but only one contains auth", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1",
                groupName: "1",
                data: {
                  authName: "auth1",
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2",
                groupName: "2",
                data: {
                  serverUrl: "https://server2",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(["operation1", "operation2"], inputs);

          assert.isUndefined(res);
        });

        it(" validate operations should return not error message when select APIs with multiple auth", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
            supportedApisFromApiSpec: [
              {
                id: "operation1",
                label: "operation1-label",
                groupName: "1",
                data: {
                  authName: "auth1",
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation2",
                label: "operation2-label",
                groupName: "2",
                data: {
                  authName: "auth2",
                  serverUrl: "https://server1",
                },
              },
              {
                id: "operation3",
                label: "operation3-label",
                groupName: "1",
                data: {
                  authName: "auth3",
                  serverUrl: "https://server1",
                },
              },
            ],
          };

          const validationSchema = question.validation as FuncValidation<string[]>;
          const res = await validationSchema.validFunc!(["operation1", "operation2"], inputs);

          assert.equal(res, undefined);
        });

        it("list operations error", async () => {
          const question = apiOperationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
          };

          let fxError: FxError | undefined = undefined;
          try {
            await question.dynamicOptions!(inputs);
          } catch (e) {
            fxError = e;
          }

          assert.isTrue(fxError !== undefined);
        });
      });

      describe("apiSpecLocationQuestion", async () => {
        it("invalid API spec location", async () => {
          const question = apiSpecLocationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
          };
          sandbox.stub(fs, "pathExists").resolves(false);

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);

          assert.isNotEmpty(res);
        });

        it("valid API spec selecting from local file with warning", async () => {
          const question = apiSpecLocationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
          };

          sandbox.stub(SpecParser.prototype, "validate").resolves({
            status: ValidationStatus.Warning,
            errors: [],
            warnings: [{ content: "warn", type: WarningType.Unknown }],
          });
          sandbox.stub(SpecParser.prototype, "list").resolves({
            APIs: [
              {
                api: "get operation1",
                server: "https://server",
                auth: {
                  name: "apiKey",
                  authScheme: {
                    type: "apiKey",
                    in: "header",
                    name: "api_key",
                  },
                },
                operationId: "getOperation1",
                isValid: true,
                reason: [],
              },
              {
                api: "get operation2",
                server: "https://server2",
                operationId: "getOperation2",
                isValid: true,
                reason: [],
              },
              {
                api: "get operation3",
                server: "https://server",
                operationId: "getOperation3",
                isValid: true,
                reason: [],
                auth: {
                  name: "authName",
                  authScheme: {
                    type: "oauth2",
                    flows: {
                      authorizationCode: {
                        authorizationUrl: "url",
                        tokenUrl: "url",
                        scopes: {},
                      },
                    },
                  },
                },
              },
              {
                api: "get operation4",
                server: "https://server",
                operationId: "getOperation4",
                isValid: true,
                reason: [],
                auth: {
                  name: "",
                  authScheme: {
                    type: "openIdConnect",
                    openIdConnectUrl: "url",
                  },
                },
              },
            ],
            allAPICount: 2,
            validAPICount: 2,
          });
          sandbox.stub(fs, "pathExists").resolves(true);

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);
          assert.deepEqual(inputs.supportedApisFromApiSpec, [
            {
              id: "get operation1",
              label: "get operation1",
              detail: "API Key authentication(In header or query)",
              groupName: "GET",
              data: {
                authName: "apiKey",
                serverUrl: "https://server",
                authType: "apiKey",
              },
            },
            {
              id: "get operation2",
              label: "get operation2",
              detail: "No authentication",
              groupName: "GET",
              data: {
                serverUrl: "https://server2",
              },
            },
            {
              id: "get operation3",
              label: "get operation3",
              detail: "OAuth(Authorization code flow)",
              groupName: "GET",
              data: {
                serverUrl: "https://server",
                authType: "oauth2",
                authName: "authName",
              },
            },
            {
              id: "get operation4",
              label: "get operation4",
              detail: "Unsupported Authorization type",
              groupName: "GET",
              data: {
                serverUrl: "https://server",
              },
            },
          ]);
          assert.isUndefined(res);
        });

        it("valid API spec of remote URL", async () => {
          const question = apiSpecLocationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
          };

          sandbox
            .stub(SpecParser.prototype, "validate")
            .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });

          sandbox.stub(SpecParser.prototype, "list").resolves({
            APIs: [
              {
                api: "get operation1",
                server: "https://server",
                auth: {
                  name: "bearerAuth",
                  authScheme: {
                    type: "http",
                    scheme: "bearer",
                  },
                },
                operationId: "getOperation1",
                isValid: true,
                reason: [],
              },

              {
                api: "get operation2",
                server: "https://server2",
                operationId: "getOperation2",
                isValid: true,
                reason: [],
              },
            ],
            allAPICount: 2,
            validAPICount: 2,
          });

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("https://www.test.com", inputs);
          assert.deepEqual(inputs.supportedApisFromApiSpec, [
            {
              id: "get operation1",
              label: "get operation1",
              detail: "API Key authentication(Bearer token authentication)",
              groupName: "GET",
              data: {
                authName: "bearerAuth",
                serverUrl: "https://server",
                authType: "apiKey",
              },
            },
            {
              id: "get operation2",
              label: "get operation2",
              detail: "No authentication",
              groupName: "GET",
              data: {
                serverUrl: "https://server2",
              },
            },
          ]);
          assert.isUndefined(res);
        });

        it("throw error if missing inputs", async () => {
          const question = apiSpecLocationQuestion();

          sandbox
            .stub(SpecParser.prototype, "validate")
            .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
          sandbox.stub(fs, "pathExists").resolves(true);

          sandbox.stub(SpecParser.prototype, "list").resolves({
            APIs: [
              {
                api: "get operation1",
                server: "https://server",
                auth: {
                  name: "api_key",
                  authScheme: {
                    name: "api_key",
                    in: "header",
                    type: "apiKey",
                  },
                },
                operationId: "getOperation1",
                isValid: true,
                reason: [],
              },
              {
                api: "get operation2",
                server: "https://server2",
                operationId: "getOperation2",
                isValid: true,
                reason: [],
              },
            ],
            allAPICount: 2,
            validAPICount: 2,
          });

          let err: Error | undefined = undefined;
          try {
            const validationSchema = question.validation as FuncValidation<string>;

            await validationSchema.validFunc!("https://www.test.com", undefined);
          } catch (e) {
            err = e as Error;
          }

          assert.isTrue(err?.message.includes("inputs is undefined"));
        });

        it("invalid api spec", async () => {
          const question = apiSpecLocationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
          };
          sandbox.stub(fs, "pathExists").resolves(true);
          sandbox.stub(SpecParser.prototype, "validate").resolves({
            status: ValidationStatus.Error,
            errors: [
              {
                type: ErrorType.SpecNotValid,
                content: "error",
              },
            ],
            warnings: [],
          });

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);

          assert.equal(res, "error");
        });

        it("invalid api spec - multiple errors", async () => {
          const question = apiSpecLocationQuestion();
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.ApiSpecLocation]: "apispec",
          };
          sandbox.stub(SpecParser.prototype, "validate").resolves({
            status: ValidationStatus.Error,
            errors: [
              {
                type: ErrorType.SpecNotValid,
                content: "error",
              },
              {
                type: ErrorType.RelativeServerUrlNotSupported,
                content: "error2",
              },
            ],
            warnings: [],
          });
          sandbox.stub(fs, "pathExists").resolves(true);

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);

          assert.equal(
            res,
            getLocalizedString(
              "core.createProjectQuestion.apiSpec.multipleValidationErrors.vscode.message"
            )
          );
        });

        it("invalid api spec - multiple errors CLI", async () => {
          const question = apiSpecLocationQuestion();
          const inputs: Inputs = {
            platform: Platform.CLI,
            [QuestionNames.ApiSpecLocation]: "apispec",
          };
          sandbox.stub(fs, "pathExists").resolves(true);
          sandbox.stub(SpecParser.prototype, "validate").resolves({
            status: ValidationStatus.Error,
            errors: [
              {
                type: ErrorType.SpecNotValid,
                content: "error",
              },
              {
                type: ErrorType.RelativeServerUrlNotSupported,
                content: "error2",
              },
            ],
            warnings: [],
          });

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);

          assert.equal(
            res,
            `error\n${getLocalizedString("core.common.RelativeServerUrlNotSupported")}`
          );
        });

        it("list operations without existing APIs", async () => {
          const question = apiSpecLocationQuestion(false);
          const inputs: Inputs = {
            platform: Platform.VSCode,
            "manifest-path": "fakePath",
          };

          sandbox
            .stub(SpecParser.prototype, "validate")
            .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
          sandbox.stub(SpecParser.prototype, "list").resolves({
            APIs: [
              {
                api: "GET /user/{userId}",
                server: "https://server",
                auth: {
                  name: "api_key",
                  authScheme: {
                    name: "api_key",
                    in: "header",
                    type: "apiKey",
                  },
                },
                operationId: "getUserById",
                isValid: true,
                reason: [],
              },
              {
                api: "GET /store/order",
                server: "https://server2",
                operationId: "getStoreOrder",
                isValid: true,
                reason: [],
              },
            ],
            allAPICount: 2,
            validAPICount: 2,
          });

          sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
          sandbox.stub(manifestUtils, "getOperationIds").returns(["getUserById"]);
          sandbox.stub(fs, "pathExists").resolves(true);

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);
          assert.deepEqual(inputs.supportedApisFromApiSpec, [
            {
              id: "GET /store/order",
              label: "GET /store/order",
              detail: "No authentication",
              groupName: "GET",
              data: {
                serverUrl: "https://server2",
              },
            },
          ]);
          assert.isUndefined(res);
        });

        it("No extra API found", async () => {
          const question = apiSpecLocationQuestion(false);
          const inputs: Inputs = {
            platform: Platform.VSCode,
            "manifest-path": "fakePath",
          };

          sandbox
            .stub(SpecParser.prototype, "validate")
            .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });

          sandbox.stub(SpecParser.prototype, "list").resolves({
            APIs: [
              {
                api: "GET /user/{userId}",
                server: "https://server",
                auth: {
                  name: "api_key",
                  authScheme: {
                    name: "api_key",
                    in: "header",
                    type: "apiKey",
                  },
                },
                operationId: "getUserById",
                isValid: true,
                reason: [],
              },
              {
                api: "GET /store/order",
                server: "https://server2",
                operationId: "getStoreOrder",
                isValid: true,
                reason: [],
              },
            ],
            allAPICount: 2,
            validAPICount: 2,
          });
          sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
          sandbox.stub(manifestUtils, "getOperationIds").returns(["getUserById", "getStoreOrder"]);
          sandbox.stub(fs, "pathExists").resolves(true);

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);
          assert.isNotNull(res);
        });

        it("list operations without existing APIs if Copilot plugin", async () => {
          const question = apiSpecLocationQuestion(false);
          const inputs: Inputs = {
            platform: Platform.VSCode,
            "manifest-path": "fakePath",
            [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
            [QuestionNames.DestinationApiSpecFilePath]: "openapi.yaml",
          };

          sandbox
            .stub(SpecParser.prototype, "validate")
            .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
          sandbox
            .stub(SpecParser.prototype, "list")
            .onFirstCall()
            .resolves({
              APIs: [
                {
                  api: "GET /user/{userId}",
                  server: "https://server",
                  operationId: "getUserById",
                  isValid: true,
                  reason: [],
                },
                {
                  api: "GET /store/order",
                  server: "https://server2",
                  operationId: "getStoreOrder",
                  isValid: true,
                  reason: [],
                },
              ],
              allAPICount: 2,
              validAPICount: 2,
            })
            .onSecondCall()
            .resolves({
              APIs: [
                {
                  api: "GET /store/order",
                  server: "https://server2",
                  operationId: "getStoreOrder",
                  isValid: true,
                  reason: [],
                },
              ],
              allAPICount: 1,
              validAPICount: 1,
            });

          sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
          sandbox
            .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
            .resolves(ok(["openapi.yaml"]));
          sandbox.stub(fs, "pathExists").resolves(true);

          const validationSchema = question.validation as FuncValidation<string>;
          const res = await validationSchema.validFunc!("file", inputs);
          assert.deepEqual(inputs.supportedApisFromApiSpec, [
            {
              data: {
                serverUrl: "https://server",
              },
              groupName: "GET",
              detail: "No authentication",
              id: "GET /user/{userId}",
              label: "GET /user/{userId}",
            },
          ]);
          assert.isUndefined(res);
        });
      });
    });

    describe("declarative copilot", () => {
      let mockedEnvRestore: RestoreFn;
      const tools = new MockTools();
      setTools(tools);
      beforeEach(() => {
        mockedEnvRestore = mockedEnv({});
      });
      afterEach(() => {
        if (mockedEnvRestore) {
          mockedEnvRestore();
        }
      });

      it("declarative copilot without plugin", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);

          await callFuncs(question, inputs);

          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 1);
            const title =
              typeof question.title === "function" ? await question.title(inputs) : question.title;
            assert.equal(
              title,
              getLocalizedString("core.createProjectQuestion.projectType.copilotExtension.title")
            );
            return ok({ type: "success", result: CapabilityOptions.declarativeAgent().id });
          } else if (question.name === QuestionNames.WithPlugin) {
            return ok({ type: "success", result: "no" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.WithPlugin,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("declarative copilot with plugin from scratch", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);

          await callFuncs(question, inputs);

          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 6);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            assert.isTrue(options.length === 1);

            return ok({ type: "success", result: CapabilityOptions.declarativeAgent().id });
          } else if (question.name === QuestionNames.WithPlugin) {
            const select = question as SingleSelectQuestion;
            const options = select.staticOptions;
            assert.isTrue(options.length === 2);
            return ok({ type: "success", result: DeclarativeCopilotTypeOptions.withPlugin().id });
          } else if (question.name === QuestionNames.ProgrammingLanguage) {
            return ok({ type: "success", result: "javascript" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.WithPlugin,
          QuestionNames.ApiPluginType,
          QuestionNames.ProgrammingLanguage,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("declarative copilot with existing plugin", async () => {
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const questions: string[] = [];
        const visitor: QuestionTreeVisitor = async (
          question: Question,
          ui: UserInteraction,
          inputs: Inputs,
          step?: number,
          totalSteps?: number
        ) => {
          questions.push(question.name);

          await callFuncs(question, inputs);

          if (question.name === QuestionNames.ProjectType) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            return ok({ type: "success", result: ProjectTypeOptions.Agent().id });
          } else if (question.name === QuestionNames.Capabilities) {
            const select = question as SingleSelectQuestion;
            const options = await select.dynamicOptions!(inputs);
            const title =
              typeof question.title === "function" ? await question.title(inputs) : question.title;
            assert.equal(
              title,
              getLocalizedString("core.createProjectQuestion.projectType.copilotExtension.title")
            );
            return ok({ type: "success", result: CapabilityOptions.declarativeAgent().id });
          } else if (question.name === QuestionNames.WithPlugin) {
            return ok({ type: "success", result: DeclarativeCopilotTypeOptions.withPlugin().id });
          } else if (question.name === QuestionNames.ApiPluginType) {
            return ok({ type: "success", result: ApiPluginStartOptions.existingPlugin().id });
          } else if (question.name === QuestionNames.PluginManifestFilePath) {
            const select = question as SingleFileQuestion;
            const title = select.title;
            assert.isNotEmpty(title);

            const defaultFolderFunc = select.defaultFolder as LocalFunc<string>;
            let defaultFolder = await defaultFolderFunc(inputs);
            assert.notEqual(defaultFolder, "./");
            defaultFolder = await defaultFolderFunc({ ...inputs, platform: Platform.CLI });
            assert.equal(defaultFolder, "./");

            sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
              ok({
                schema_version: "v2.0",
                name_for_human: "test",
                runtimes: [
                  {
                    type: "OpenApi",
                    spec: {
                      url: "test.json",
                    },
                  },
                ],
              } as any)
            );
            const validationFunc = question.validation as FuncValidation<string>;
            const validationRes = await validationFunc.validFunc!("", inputs);
            assert.isUndefined(validationRes);

            return ok({ type: "success", result: "c://testFolder/test.json" });
          } else if (question.name === QuestionNames.PluginOpenApiSpecFilePath) {
            const select = question as SingleFileQuestion;
            const title = select.title;
            assert.isNotEmpty(title);

            const defaultFolderFunc = select.defaultFolder as LocalFunc<string>;
            let defaultFolder = await defaultFolderFunc(inputs);
            assert.isTrue(defaultFolder.endsWith("testFolder"));
            defaultFolder = await defaultFolderFunc({ ...inputs, platform: Platform.CLI });

            assert.equal(defaultFolder, "./");

            sandbox.stub(SpecParser.prototype, "validate").resolves({
              status: ValidationStatus.Valid,
              errors: [],
              warnings: [],
            });
            const validationFunc = question.validation as FuncValidation<string>;
            const validationRes = await validationFunc.validFunc!("test.json", inputs);
            assert.isUndefined(validationRes);

            return ok({ type: "success", result: "test.json" });
          } else if (question.name === QuestionNames.AppName) {
            return ok({ type: "success", result: "test001" });
          } else if (question.name === QuestionNames.Folder) {
            return ok({ type: "success", result: "./" });
          }
          return ok({ type: "success", result: undefined });
        };
        await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
        assert.deepEqual(questions, [
          QuestionNames.ProjectType,
          QuestionNames.Capabilities,
          QuestionNames.WithPlugin,
          QuestionNames.ApiPluginType,
          QuestionNames.PluginManifestFilePath,
          QuestionNames.PluginOpenApiSpecFilePath,
          QuestionNames.Folder,
          QuestionNames.AppName,
        ]);
      });

      it("pluginManifestQuestion: Invalid due to read manifest error ", async () => {
        const question = pluginManifestQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox
          .stub(pluginManifestUtils, "readPluginManifestFile")
          .resolves(err(new UserError("source", "name", "fakeError", "fakeError")));
        const validationFunc = question.validation as FuncValidation<string>;
        const validationRes = await validationFunc.validFunc!("", inputs);
        assert.equal(validationRes, "fakeError");
      });

      it("pluginManifestQuestion: Invalid due to missing runtime", async () => {
        const question = pluginManifestQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
          ok({
            schema_version: "v2.0",
            name_for_human: "test",
            runtimes: [],
          } as any)
        );
        const validationFunc = question.validation as FuncValidation<string>;
        const validationRes = await validationFunc.validFunc!("", inputs);
        assert.isTrue(validationRes?.includes("OpenApi"));
      });

      it("pluginManifestQuestion: default plugin manifest", async () => {
        const question = pluginManifestQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
          projectPath: "./",
        };
        sandbox.stub(fs, "existsSync").returns(true);
        const defaultFile = question.default as LocalFunc<string>;
        const res = await defaultFile(inputs);
        assert.exists(res);
      });

      it("pluginManifestQuestion: projectpath not exists", async () => {
        const question = pluginManifestQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox.stub(fs, "existsSync").returns(true);
        const defaultFile = question.default as LocalFunc<string>;
        const res = await defaultFile(inputs);
        assert.isUndefined(res);
      });

      it("pluginManifestQuestion: default file not exists", async () => {
        const question = pluginManifestQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
          projectPath: "./",
        };
        sandbox.stub(fs, "existsSync").returns(false);
        const defaultFile = question.default as LocalFunc<string>;
        const res = await defaultFile(inputs);
        assert.isUndefined(res);
      });

      it("pluginApiSpecQuestion: invalid file format", async () => {
        const question = pluginApiSpecQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };

        const validationFunc = question.validation as FuncValidation<string>;
        const validationRes = await validationFunc.validFunc!("test.txt", inputs);
        assert.isTrue(validationRes?.includes("json, yml, yaml"));
      });

      it("pluginApiSpecQuestion: invalid spec ", async () => {
        const question = pluginApiSpecQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox.stub(SpecParser.prototype, "validate").resolves({
          status: ValidationStatus.Error,
          errors: [
            {
              type: ErrorType.SpecNotValid,
              content: "invalidFile",
            },
          ],
          warnings: [],
        });
        const validationFunc = question.validation as FuncValidation<string>;
        const validationRes = await validationFunc.validFunc!("test.json", inputs);
        assert.equal(validationRes, "invalidFile");
      });
    });
  });

  describe("createProjectQuestionNode if chatParticipant is enabled", async () => {
    const ui = new MockUserInteraction();
    let mockedEnvRestore: RestoreFn = () => {};

    beforeEach(() => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.SampleConfigBranch]: "dev",
        [FeatureFlagName.ChatParticipantUIEntries]: "true",
      });
    });
    afterEach(() => {
      mockedEnvRestore();
    });

    it("notification bot", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 7);
          assert.equal(
            getLocalizedString("core.createProjectQuestion.projectType.createGroup.aiAgent"),
            (options as OptionItem[])[0].groupName
          );
          return ok({ type: "success", result: ProjectTypeOptions.bot().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: CapabilityOptions.notificationBot().id });
        } else if (question.name === QuestionNames.BotTrigger) {
          return ok({ type: "success", result: NotificationTriggerOptions.appService().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          return ok({ type: "success", result: "javascript" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.BotTrigger,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });

    it("chat with Copilot Chat", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);

        await callFuncs(question, inputs);

        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 6);
          return ok({ type: "success", result: ProjectTypeOptions.startWithGithubCopilot().id });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [QuestionNames.ProjectType]);
    });
  });
  describe("getLanguageOptions", () => {
    let mockedEnvRestore: RestoreFn = () => {};

    afterEach(() => {
      mockedEnvRestore();
    });

    it("dotnet for VS", async () => {
      const options = getLanguageOptions({
        platform: Platform.VS,
        runtime: RuntimeOptions.DotNet().id,
      });
      assert.isTrue(options.length === 1 && options[0].id === "csharp");
    });

    it("dotnet when TEAMSFX_CLI_DOTNET", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "true" });
      const options = getLanguageOptions({
        platform: Platform.CLI,
        runtime: RuntimeOptions.DotNet().id,
      });
      assert.isTrue(options.length === 1 && options[0].id === "csharp");
    });
    it("SPFx", async () => {
      const options = getLanguageOptions({
        platform: Platform.VSCode,
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.SPFxTab().id,
      });
      assert.isTrue(options.length === 1 && options[0].id === "typescript");
    });
    it("other", async () => {
      const options = getLanguageOptions({
        platform: Platform.VSCode,
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.basicBot().id,
      });
      assert.isTrue(options.length === 2);
    });
    it("should return python for ProgrammingLanguage.PY", () => {
      const output = convertToLangKey(ProgrammingLanguage.PY);
      assert.isTrue(output == "python");
    });
    it("should return expected 3 language options for custom copilot basic python", () => {
      const options = getLanguageOptions({
        platform: Platform.VSCode,
        [QuestionNames.ProjectType]: ProjectTypeOptions.customCopilot().id,
        [QuestionNames.Capabilities]: CapabilityOptions.customCopilotBasic().id,
      });
      assert.isTrue(options.length === 3); // js, ts, python
    });
  });

  describe("appNameQuestion", () => {
    const question = appNameQuestion();
    const validFunc = (question.validation as FuncValidation<string>).validFunc;
    it("happy path", async () => {
      const inputs: Inputs = { platform: Platform.VSCode, folder: "./" };
      const appName = "1234";
      let validRes = await validFunc(appName, inputs);
      assert.isTrue(validRes === getLocalizedString("core.QuestionAppName.validation.pattern"));
      sandbox.stub<any, any>(fs, "pathExists").resolves(true);
      inputs.appName = randomAppName();
      inputs.folder = "./";
      validRes = await validFunc(inputs.appName, inputs);
      const expected = getLocalizedString(
        "core.QuestionAppName.validation.pathExist",
        path.resolve(inputs.folder, inputs.appName)
      );
      assert.equal(validRes, expected);
      sandbox.restore();
      sandbox.stub<any, any>(fs, "pathExists").resolves(false);
      validRes = await validFunc(inputs.appName, inputs);
      assert.isTrue(validRes === undefined);
    });

    it("app name has 25 length - VSC", async () => {
      const mockedUI = new MockedUserInteraction();
      sandbox.stub(utils, "createContext").returns({
        userInteraction: mockedUI,
      } as Context);
      const showMessageStub = sandbox.stub(mockedUI, "showMessage");

      const input = "abcdefghijklmnopqrstuvwxy";
      await validFunc(input, { platform: Platform.VSCode });

      assert.isTrue(showMessageStub.calledOnce);
    });

    it("app name has 25 length - VS", async () => {
      const mockedLogProvider = new MockedLogProvider();
      sandbox.stub(utils, "createContext").returns({
        logProvider: mockedLogProvider as LogProvider,
      } as Context);
      const warningStub = sandbox.stub(mockedLogProvider, "warning");

      const input = "abcdefghijklmnopqrstuvwxy";
      await validFunc(input, { platform: Platform.VS });

      assert.isTrue(warningStub.calledOnce);

      await validFunc(input);

      assert.isTrue(warningStub.calledTwice);
    });

    it("app name exceed maxlength of 30", async () => {
      const input = "SurveyMonkeyWebhookNotification";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.maxlength"));
    });

    it("app name with only letters", async () => {
      const input = "app";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name starting with digit", async () => {
      const input = "123app";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name count of alphanumerics less than 2", async () => {
      const input = "a..(";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing dot", async () => {
      const input = "app.123";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing hyphen", async () => {
      const input = "app-123";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing multiple special characters", async () => {
      const input = "a..(1";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing space", async () => {
      const input = "app 123";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing dot at the end - wrong pattern", async () => {
      const input = "app.app.";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing space at the end - wrong pattern", async () => {
      const input = "app123 ";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing invalid control code", async () => {
      const input = "a\u0001a";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing invalid character", async () => {
      const input = "app<>123";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("invalid app name containing &", async () => {
      const input = "app&123";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });
  });

  describe("CapabilityOptions", () => {
    let mockedEnvRestore: RestoreFn = () => {};
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({});
    });
    afterEach(() => {
      mockedEnvRestore();
    });
    it("has 3 options in message extension type", () => {
      // Act
      const options = CapabilityOptions.mes();
      // Assert
      assert.equal(options.length, 3);
      assert.deepEqual(options, [
        CapabilityOptions.m365SearchMe(),
        CapabilityOptions.collectFormMe(),
        CapabilityOptions.linkUnfurling(),
      ]);
    });
    it("cli non-interactive", () => {
      const question = capabilityQuestion();
      const options = question.dynamicOptions!({ platform: Platform.CLI, nonInteractive: true });
      assert.deepEqual(
        options,
        CapabilityOptions.all({ platform: Platform.CLI, nonInteractive: true })
      );
    });
    it("vs non-interactive", () => {
      const question = capabilityQuestion();
      const options = question.dynamicOptions!({ platform: Platform.VS });
      assert.deepEqual(options, CapabilityOptions.dotnetCaps({ platform: Platform.VS }));
    });

    it("templates for TDP integration", () => {
      mockedEnvRestore();
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.TdpTemplateCliTest]: "true",
      });
      const question = capabilityQuestion();
      const options = question.dynamicOptions!({
        platform: Platform.CLI,
        nonInteractive: true,
      }) as OptionItem[];
      assert.isTrue(options.findIndex((o: OptionItem) => o.id === CapabilityOptions.me().id) > -1);
      assert.isTrue(
        options.findIndex((o: OptionItem) => o.id === CapabilityOptions.botAndMe().id) > -1
      );
      assert.isTrue(
        options.findIndex((o: OptionItem) => o.id === CapabilityOptions.nonSsoTabAndBot().id) > -1
      );
    });

    it("templates for TDP integration dotnet", () => {
      mockedEnvRestore();
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.TdpTemplateCliTest]: "true",
        [FeatureFlagName.CLIDotNet]: "true",
      });
      const question = capabilityQuestion();
      const options = question.dynamicOptions!({
        platform: Platform.CLI,
        nonInteractive: true,
        runtime: "dotnet",
      }) as OptionItem[];
      assert.isTrue(options.findIndex((o: OptionItem) => o.id === CapabilityOptions.me().id) > -1);
      assert.isTrue(
        options.findIndex((o: OptionItem) => o.id === CapabilityOptions.botAndMe().id) < 0
      );
      assert.isTrue(
        options.findIndex((o: OptionItem) => o.id === CapabilityOptions.nonSsoTabAndBot().id) < 0
      );
    });

    describe("officeAddinCapabilities()", () => {
      it("should return correct capabilities for outlook addin", () => {
        const capabilities = CapabilityOptions.officeAddinCapabilities(
          ProjectTypeOptions.outlookAddin().id
        );
        assert.equal(capabilities.length, 2);
      });
      it("should return correct capabilities for office addin", () => {
        const capabilities = CapabilityOptions.officeAddinCapabilities(
          ProjectTypeOptions.officeMetaOS().id
        );
        assert.equal(capabilities.length, 2);
      });
    });
  });

  describe("ME copilot plugin template only", () => {
    const ui = new MockUserInteraction();
    let mockedEnvRestore: RestoreFn;
    const tools = new MockTools();
    setTools(tools);
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.ChatParticipantUIEntries]: "false",
      });
    });

    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("traverse in vscode ME Copilot Plugin", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const questions: string[] = [];
      const visitor: QuestionTreeVisitor = async (
        question: Question,
        ui: UserInteraction,
        inputs: Inputs,
        step?: number,
        totalSteps?: number
      ) => {
        questions.push(question.name);
        await callFuncs(question, inputs);
        if (question.name === QuestionNames.ProjectType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          return ok({ type: "success", result: ProjectTypeOptions.me().id });
        } else if (question.name === QuestionNames.Capabilities) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          assert.deepEqual(options, CapabilityOptions.mes());
          return ok({ type: "success", result: CapabilityOptions.m365SearchMe().id });
        } else if (question.name === QuestionNames.MeArchitectureType) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 3);
          return ok({ type: "success", result: MeArchitectureOptions.botMe().id });
        } else if (question.name === QuestionNames.ProgrammingLanguage) {
          const select = question as SingleSelectQuestion;
          const options = await select.dynamicOptions!(inputs);
          assert.isTrue(options.length === 2);
          return ok({ type: "success", result: "typescript" });
        } else if (question.name === QuestionNames.Folder) {
          return ok({ type: "success", result: "./" });
        } else if (question.name === QuestionNames.AppName) {
          return ok({ type: "success", result: "test001" });
        }
        return ok({ type: "success", result: undefined });
      };
      await traverse(createProjectQuestionNode(), inputs, ui, undefined, visitor);
      assert.deepEqual(questions, [
        QuestionNames.ProjectType,
        QuestionNames.Capabilities,
        QuestionNames.MeArchitectureType,
        QuestionNames.ProgrammingLanguage,
        QuestionNames.Folder,
        QuestionNames.AppName,
      ]);
    });
  });
  describe("programmingLanguageQuestion", () => {
    const question = programmingLanguageQuestion();
    it("SPFxTab", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Capabilities]: CapabilityOptions.SPFxTab().id,
      };
      if (
        question.dynamicOptions &&
        question.placeholder &&
        typeof question.placeholder === "function"
      ) {
        const options = question.dynamicOptions(inputs);
        assert.deepEqual([{ id: "typescript", label: "TypeScript" }], options);
        const placeholder = question.placeholder(inputs);
        assert.equal("SPFx is currently supporting TypeScript only.", placeholder);
      }

      languageAssert({
        platform: Platform.VSCode,
        [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
      });
      languageAssert({
        platform: Platform.VSCode,
        [QuestionNames.Capabilities]: CapabilityOptions.basicBot().id,
      });
      languageAssert({
        platform: Platform.VSCode,
        [QuestionNames.Capabilities]: CapabilityOptions.me().id,
      });

      function languageAssert(inputs: Inputs) {
        if (
          question.dynamicOptions &&
          question.placeholder &&
          typeof question.placeholder === "function"
        ) {
          const options = question.dynamicOptions(inputs);
          assert.deepEqual(
            [
              { id: "javascript", label: "JavaScript" },
              { id: "typescript", label: "TypeScript" },
            ],
            options
          );
          const placeholder = question.placeholder(inputs);
          assert.equal("Select a programming language", placeholder);
        }
      }
    });
  });

  describe("folderQuestion", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("should find taskpane template", () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
      };
      const question = folderQuestion() as any;
      const title = question.title(inputs);
      const defaultV = question.default(inputs);
      assert.equal(title, "Directory where the project folder will be created in");
      assert.equal(defaultV, "./");
    });
  });

  describe("getSolutionName", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJson").resolves({
        "@microsoft/generator-sharepoint": {
          solutionName: "testSolutionName",
        },
      });
      const res = await getSolutionName("");
      assert.equal(res, "testSolutionName");
    });

    it("FileNotFoundError", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      try {
        await getSolutionName(".");
        assert.fail("should throw");
      } catch (e) {
        assert.isTrue(e instanceof FileNotFoundError);
      }
    });

    it("undefined", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJson").resolves({});
      const res = await getSolutionName("");
      assert.isUndefined(res);
    });
  });

  describe("api plugin auth question", () => {
    const ui = new MockUserInteraction();
    let mockedEnvRestore: RestoreFn;
    const tools = new MockTools();
    setTools(tools);
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({});
    });

    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("api message extension", async () => {
      const question = apiAuthQuestion();
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.MeArchitectureType] = MeArchitectureOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.none(),
          ApiAuthOptions.bearerToken(),
          ApiAuthOptions.microsoftEntra(),
        ]);
      }
    });

    it("api plugin from scratch with auth enabled", async () => {
      const question = apiAuthQuestion();
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.none(),
          ApiAuthOptions.apiKey(),
          ApiAuthOptions.microsoftEntra(),
          ApiAuthOptions.oauth(),
        ]);
      }
    });

    it("api plugin from add action with auth enabled", async () => {
      const question = apiAuthQuestion(true);
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.apiKey(),
          ApiAuthOptions.microsoftEntra(),
          ApiAuthOptions.oauth(),
        ]);
      }
    });
  });
  describe("api plugin auth question (AAD disabled)", () => {
    let mockedEnvRestore: RestoreFn;
    const tools = new MockTools();
    setTools(tools);
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({});
    });

    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("api plugin from scratch without AAD enabled", async () => {
      const question = apiAuthQuestion();
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.none(),
          ApiAuthOptions.apiKey(),
          ApiAuthOptions.microsoftEntra(),
          ApiAuthOptions.oauth(),
        ]);
      }
    });
  });
});
