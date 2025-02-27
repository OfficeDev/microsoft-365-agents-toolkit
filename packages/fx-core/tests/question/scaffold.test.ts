// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ConditionFunc,
  Inputs,
  LocalFunc,
  Platform,
  SingleSelectQuestion,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { featureFlagManager } from "../../src/common/featureFlags";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { Bot } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/bot";
import { MessagingExtension } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/messagingExtension";
import { StaticTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/staticTab";
import { TemplateNames } from "../../src/component/generator/templates/templateNames";
import { ProgrammingLanguage, QuestionNames } from "../../src/question/constants";
import { scaffoldQuestionForVS } from "../../src/question/scaffold/vs/createRootNode";
import {
  ApiPluginStartOptions,
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  DACapabilityOptions,
  MeCapabilityOptions,
  OfficeAddinCapabilityOptions,
  TabCapabilityOptions,
} from "../../src/question/scaffold/vsc/CapabilityOptions";
import { ProjectTypeOptions } from "../../src/question/scaffold/vsc/ProjectTypeOptions";
import {
  createFromTdpNode,
  getTemplateName,
  isTdpTemplate,
} from "../../src/question/scaffold/vsc/createFromTdpNode";
import {
  folderAndAppNameCondition,
  getProjectTypeByCapability,
  languageNode,
  scaffoldQuestionForVSCode,
} from "../../src/question/scaffold/vsc/createRootNode";
import { officeAddinProjectTypeNode } from "../../src/question/scaffold/vsc/officeAddinProjectTypeNode";
import { apiSpecNode } from "../../src/question/scaffold/vsc/teamsProjectTypeNode";
import { TdpCapabilityOptions } from "../../build/question/scaffold/vsc/createFromTdpNode";

describe("vsc", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("scaffoldQuestionForVSCode", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const root = scaffoldQuestionForVSCode();
    assert.isDefined(root);
  });
  it("scaffoldQuestionForVSCode", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const root = scaffoldQuestionForVSCode();
    assert.isDefined(root);
  });
  it("createFromTdpNode", () => {
    const root = createFromTdpNode();
    assert.isDefined(root);
  });
});

describe("vs", () => {
  it("scaffoldQuestionForVS", () => {
    const root = scaffoldQuestionForVS();
    assert.isDefined(root);
  });
});

describe("getTemplateName", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  const validBot: Bot = {
    botId: "botId",
    isNotificationOnly: false,
    needsChannelSelector: false,
    personalCommands: [{ title: "title", description: "description" }],
    supportsFiles: false,
    supportsCalling: false,
    supportsVideo: false,
    teamCommands: [{ title: "title", description: "description" }],
    groupChatCommands: [{ title: "title", description: "description" }],
    scopes: ["scope"],
  };

  const validStaticTab: StaticTab = {
    objectId: "objId",
    entityId: "entityId",
    name: "tab",
    contentUrl: "https://url",
    websiteUrl: "https:/url",
    scopes: [],
    context: [],
  };

  const validMessagingExtension: MessagingExtension = {
    objectId: "objId",
    botId: "botId",
    canUpdateConfiguration: true,
    commands: [],
    messageHandlers: [],
  };

  it("return TabNonSsoAndDefaultBot", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      staticTabs: [validStaticTab],
      messagingExtensions: [validMessagingExtension],
    };

    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: appDefinition,
    };

    const res = getTemplateName(inputs);
    assert.equal(res, TemplateNames.TabAndDefaultBot);
  });

  it("return TabNonSso", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      staticTabs: [validStaticTab],
    };

    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: appDefinition,
    };

    const res = getTemplateName(inputs);
    assert.equal(res, TemplateNames.Tab);
  });

  it("return DefaultBotAndMessageExtension", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      bots: [validBot],
      messagingExtensions: [validMessagingExtension],
    };

    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: appDefinition,
    };

    const res = getTemplateName(inputs);
    assert.equal(res, TemplateNames.BotAndMessageExtension);
  });

  it("return MessageExtension", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      messagingExtensions: [validMessagingExtension],
    };

    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: appDefinition,
    };

    const res = getTemplateName(inputs);
    assert.equal(res, TemplateNames.MessageExtension);
  });

  it("return bot", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      bots: [validBot],
    };

    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: appDefinition,
    };

    const res = getTemplateName(inputs);
    assert.equal(res, TemplateNames.DefaultBot);
  });

  it("return undefined", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
    };

    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: appDefinition,
    };

    const res = getTemplateName(inputs);
    assert.isUndefined(res);
  });

  it("tdp cli test", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const inputs: Inputs = {
      platform: Platform.CLI,
      nonInteractive: true,
      [QuestionNames.Capabilities]: TdpCapabilityOptions.me().id,
    };
    const res = getTemplateName(inputs);
    assert.equal(res, TemplateNames.MessageExtension);
  });

  it("isTdpTemplate", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const inputs: Inputs = {
      platform: Platform.CLI,
      nonInteractive: true,
      [QuestionNames.Capabilities]: TdpCapabilityOptions.me().id,
    };
    const res = isTdpTemplate(inputs);
    assert.isTrue(res);
  });
});

describe("m365ProjectTypeNode", () => {
  it("apiSpecNode", () => {
    const node = apiSpecNode({ equals: "a" });
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const condition = node.children?.[1].condition as ConditionFunc;
    const res = condition?.(inputs);
    assert.isTrue(res);
  });
});

describe("ProjectTypeOptions", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("officeMetaOS - VSC", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const option = ProjectTypeOptions.officeAddin(Platform.VSCode);
    assert.equal(option.id, ProjectTypeOptions.officeMetaOSOptionId);
  });
  it("officeMetaOS - CLI", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const option = ProjectTypeOptions.officeAddin(Platform.CLI);
    assert.equal(option.id, ProjectTypeOptions.officeMetaOSOptionId);
  });
  it("outlookAddin - VSC", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const option = ProjectTypeOptions.officeAddin(Platform.VSCode);
    assert.equal(option.id, ProjectTypeOptions.outlookAddinOptionId);
  });
  it("outlookAddin - CLI", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const option = ProjectTypeOptions.officeAddin(Platform.CLI);
    assert.equal(option.id, ProjectTypeOptions.outlookAddinOptionId);
  });
});

describe("officeAddinProjectTypeNode", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("wxpAddinProjectTypeNode", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const node = officeAddinProjectTypeNode();
    assert.deepEqual(node.condition, {
      equals: ProjectTypeOptions.officeMetaOSOptionId,
    });
  });
  it("outlookAddinProjectTypeNode", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const node = officeAddinProjectTypeNode();
    assert.deepEqual(node.condition, {
      equals: ProjectTypeOptions.outlookAddinOptionId,
    });
  });
});

describe("languageNode", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("csharp", () => {
    const node = languageNode();
    const condition = node.condition as ConditionFunc;
    const inputs: Inputs = {
      platform: Platform.VS,
      [QuestionNames.TemplateName]: TemplateNames.SsoTabSSR,
    };
    const res = condition(inputs);
    assert.isTrue(res);
    const question = node.data as SingleSelectQuestion;
    const options = question.dynamicOptions?.(inputs);
    assert.deepEqual(options, [{ id: ProgrammingLanguage.CSharp, label: "C#" }]);
    const defaultFunc = question.default as LocalFunc<string | undefined>;
    const defaultOptionId = defaultFunc ? defaultFunc(inputs) : undefined;
    assert.equal(defaultOptionId, ProgrammingLanguage.CSharp);
  });
  it("common", () => {
    const node = languageNode();
    const condition = node.condition as ConditionFunc;
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.TemplateName]: TemplateNames.BasicGpt,
    };
    const res = condition(inputs);
    assert.isTrue(res);
    const options = (node.data as SingleSelectQuestion).dynamicOptions?.(inputs);
    assert.deepEqual(options, [{ id: ProgrammingLanguage.Common, label: "None" }]);
  });
});

describe("folderAndAppNameCondition", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("ApiPluginManifestPath", () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.ApiPluginManifestPath]: "test",
    };
    const res = folderAndAppNameCondition(inputs);
    assert.isTrue(res);
  });
  it("false", () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      [QuestionNames.ProjectType]: ProjectTypeOptions.copilotAgentOptionId,
    };
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const res = folderAndAppNameCondition(inputs);
    assert.isFalse(res);
  });
});

describe("getProjectTypeByCapability", () => {
  it("DA", () => {
    const type = getProjectTypeByCapability(DACapabilityOptions.declarativeAgent().id);
    assert.equal(type, ProjectTypeOptions.copilotAgentOptionId);
  });
  it("CEA", () => {
    const type = getProjectTypeByCapability(CustomCopilotCapabilityOptions.customCopilotRag().id);
    assert.equal(type, ProjectTypeOptions.customCopilotOptionId);
  });
  it("Bot", () => {
    const type = getProjectTypeByCapability(BotCapabilityOptions.basicBot().id);
    assert.equal(type, ProjectTypeOptions.botOptionId);
  });
  it("Tab", () => {
    const type = getProjectTypeByCapability(TabCapabilityOptions.nonSsoTab().id);
    assert.equal(type, ProjectTypeOptions.tabOptionId);
  });
  it("ME", () => {
    const type = getProjectTypeByCapability(MeCapabilityOptions.m365SearchMe().id);
    assert.equal(type, ProjectTypeOptions.meOptionId);
  });
  it("WXP", () => {
    const type = getProjectTypeByCapability(OfficeAddinCapabilityOptions.wxpTaskPane().id);
    assert.equal(type, ProjectTypeOptions.officeMetaOSOptionId);
  });
  it("Outlook", () => {
    const type = getProjectTypeByCapability(OfficeAddinCapabilityOptions.outlookTaskPane().id);
    assert.equal(type, ProjectTypeOptions.outlookAddinOptionId);
  });
});
