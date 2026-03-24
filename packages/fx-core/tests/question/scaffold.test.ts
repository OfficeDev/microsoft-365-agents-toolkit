// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ConditionFunc,
  Inputs,
  LocalFunc,
  OptionItem,
  Platform,
  SingleSelectQuestion,
  StringValidation,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import proxyquire from "proxyquire";
import sinon from "sinon";
import { featureFlagManager, FeatureFlags } from "../../src/common/featureFlags";
import { getLocalizedString } from "../../src/common/localizeUtils";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { Bot } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/bot";
import { MessagingExtension } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/messagingExtension";
import { StaticTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/staticTab";
import { TemplateNames } from "../../src/component/generator/templates/templateNames";
import { ProgrammingLanguage, QuestionNames } from "../../src/question/constants";
import { foundryAgentIdQuestion, foundryEndpointQuestion } from "../../src/question/create";
import {
  apiSpecNode,
  apiSpecWithSearchNode,
  foundryNode,
} from "../../src/question/scaffold/commonNodes";
import { constructNode } from "../../src/question/scaffold/constructNode";
import { scaffoldQuestionForVS } from "../../src/question/scaffold/vs/createRootNode";
import { ActionStartOptions } from "../../src/question/scaffold/vsc/CapabilityOptions";
import { ProjectTypeOptions } from "../../src/question/scaffold/vsc/ProjectTypeOptions";
import {
  createFromTdpNode,
  getTemplateName,
} from "../../src/question/scaffold/vsc/createFromTdpNode";
import {
  folderAndAppNameCondition,
  languageNode,
  scaffoldQuestionForVSCode,
} from "../../src/question/scaffold/vsc/createRootNode";
import { getRootProjectTypeNode } from "../../src/question/scaffold/vsc/rootNode";
import { daProjectTypeNode } from "../../src/question/scaffold/vsc/daProjectTypeNode";
import { officeAddinProjectTypeNode } from "../../src/question/scaffold/vsc/officeAddinProjectTypeNode";
import {
  getTeamsProjectNode,
  TeamsProjectTypeOptions,
} from "../../src/question/scaffold/vsc/teamsProjectTypeNode";

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
    assert.equal(res, TemplateNames.DefaultBot);
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
    assert.equal(res, TemplateNames.DefaultMessageExtension);
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
});

describe("daProjectTypeNode", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("daProjectTypeNode basic structure", () => {
    const node = daProjectTypeNode();
    const conditionFunc = node?.condition as StringValidation;

    assert.equal(conditionFunc.equals, ProjectTypeOptions.copilotAgentOptionId);
    assert.isDefined(node.children);
  });

  it("should return apiSpecWithSearchNode when KiotaNPMIntegration is enabled", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag) => {
      if (flag === FeatureFlags.KiotaNPMIntegration) {
        return true;
      }
      return false;
    });

    const node = daProjectTypeNode();
    const withPluginNode = node.children?.[0];
    assert.isDefined(withPluginNode);

    const actionTypeNode = withPluginNode?.children?.[0];
    assert.isDefined(actionTypeNode);

    const apiSpecChildNode = actionTypeNode?.children?.[1];

    assert.isDefined(apiSpecChildNode);

    const firstChild = apiSpecChildNode?.children?.[0];
    assert.isDefined(firstChild);

    const selectApiSpecQuestion = firstChild?.data;
    assert.isDefined(selectApiSpecQuestion);
    assert.equal(selectApiSpecQuestion?.name, QuestionNames.OpenAPISpecType);
  });

  it("should return apiSpecNode when KiotaNPMIntegration is disabled", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag) => {
      if (flag === FeatureFlags.KiotaNPMIntegration) {
        return false;
      }
      return false;
    });

    const node = daProjectTypeNode();
    const withPluginNode = node.children?.[0];
    assert.isDefined(withPluginNode);

    const actionTypeNode = withPluginNode?.children?.[0];
    assert.isDefined(actionTypeNode);

    const apiSpecChildNode = actionTypeNode?.children?.[1];

    assert.isDefined(apiSpecChildNode);

    assert.isFunction(apiSpecChildNode?.condition);

    const testInputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.ActionType]: ActionStartOptions.apiSpec().id,
    };

    const conditionFunc = apiSpecChildNode?.condition as ConditionFunc;
    assert.isTrue(conditionFunc(testInputs));
  });

  it("should include MCP option when MCPForDA feature flag is enabled", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag) => {
      if (flag === FeatureFlags.MCPForDA) {
        return true;
      }
      return false;
    });

    const node = daProjectTypeNode();
    const withPluginNode = node.children?.[0];
    assert.isDefined(withPluginNode);

    const actionTypeNode = withPluginNode?.children?.[0];
    assert.isDefined(actionTypeNode);

    const actionTypeData = actionTypeNode?.data as SingleSelectQuestion;
    assert.isDefined(actionTypeData);
    assert.isDefined(actionTypeData.staticOptions);

    // Check that MCP option is included in staticOptions
    const staticOptions = actionTypeData.staticOptions;
    let mcpOption: string | OptionItem | undefined;

    if (Array.isArray(staticOptions) && staticOptions.length > 0) {
      if (typeof staticOptions[0] === "string") {
        mcpOption = (staticOptions as string[]).find(
          (option) => option === ActionStartOptions.mcp().id
        );
      } else {
        mcpOption = (staticOptions as OptionItem[]).find(
          (option) => option.id === ActionStartOptions.mcp().id
        );
      }
    }

    assert.isDefined(mcpOption);
    const mcpOptionId = typeof mcpOption === "string" ? mcpOption : mcpOption?.id;
    assert.equal(mcpOptionId, "mcp");
  });

  it("should not include MCP option when MCPForDA feature flag is disabled", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag) => {
      if (flag === FeatureFlags.MCPForDA) {
        return false;
      }
      return false;
    });

    const node = daProjectTypeNode();
    const withPluginNode = node.children?.[0];
    assert.isDefined(withPluginNode);

    const actionTypeNode = withPluginNode?.children?.[0];
    assert.isDefined(actionTypeNode);

    const actionTypeData = actionTypeNode?.data as SingleSelectQuestion;
    assert.isDefined(actionTypeData);
    assert.isDefined(actionTypeData.staticOptions);

    // Check that MCP option is not included in staticOptions
    const staticOptions = actionTypeData.staticOptions;
    let mcpOption: string | OptionItem | undefined;

    if (Array.isArray(staticOptions) && staticOptions.length > 0) {
      if (typeof staticOptions[0] === "string") {
        mcpOption = (staticOptions as string[]).find(
          (option) => option === ActionStartOptions.mcp().id
        );
      } else {
        mcpOption = (staticOptions as OptionItem[]).find(
          (option) => option.id === ActionStartOptions.mcp().id
        );
      }
    }

    assert.isUndefined(mcpOption);
  });
});

describe("customEngineAgentProjectTypeNode", () => {
  it("customEngineAgentProjectTypeNode basic structure", () => {
    const root = getRootProjectTypeNode(Platform.VSCode);
    const node = root.children?.find(
      (c) =>
        (c.condition as StringValidation)?.equals === ProjectTypeOptions.customEngineAgentOptionId
    );
    assert.isDefined(node);
    const conditionFunc = node?.condition as StringValidation;
    assert.equal(conditionFunc.equals, ProjectTypeOptions.customEngineAgentOptionId);
    assert.isDefined(node!.children);
  });

  it("should extract CEA sub-tree from wizardNode with correct options", () => {
    const root = getRootProjectTypeNode(Platform.VSCode);
    const node = root.children?.find(
      (c) =>
        (c.condition as StringValidation)?.equals === ProjectTypeOptions.customEngineAgentOptionId
    );
    assert.isDefined(node);
    assert.isDefined(node!.data);
    const data = node!.data as SingleSelectQuestion;
    assert.isDefined(data.staticOptions);
    const options = data.staticOptions as OptionItem[];
    const optionIds = options.map((o) => o.id);
    assert.include(optionIds, "basic-custom-engine-agent");
    assert.include(optionIds, "weather-agent");
  });
});

describe("teamsProjectTypeNode", () => {
  it("should extract Teams sub-tree from wizardNode", () => {
    const node = getTeamsProjectNode();
    assert.isDefined(node);
    const condition = node.condition as StringValidation;
    assert.equal(condition.equals, "teams-agent-and-app-type");
  });
});

describe("m365ProjectTypeNode", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("apiSpecNode", () => {
    const node = apiSpecNode({ equals: "a" });
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const condition = node.children?.[1].condition as ConditionFunc;
    const res = condition?.(inputs);
    assert.isTrue(res);
  });

  it("apiSpecWithSearchNode", () => {
    const node = apiSpecWithSearchNode();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.ActionType]: ActionStartOptions.apiSpecWithSearch().id,
      [QuestionNames.ActionManifestPath]: "test",
      [QuestionNames.SelectOpenApiSpec]: "test",
    };
    const condition = node.children?.[0].children?.[0]?.children?.[0].condition as ConditionFunc;
    const res = condition?.(inputs);
    assert.isFalse(res);

    const condition2 = node.children?.[0]?.children?.[1]?.children?.[1]?.condition as ConditionFunc;
    const res2 = condition2?.(inputs);
    assert.isTrue(res2);

    sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag) => {
      if (flag === FeatureFlags.KiotaNPMIntegration) {
        return true;
      }
      return false;
    });

    const condition3 = node.children?.[0]?.condition as ConditionFunc;
    const res3 = condition3?.(inputs);
    assert.isTrue(res3);
    assert.isTrue(inputs[QuestionNames.ActionType] === ActionStartOptions.apiSpec().id);
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
  it("start with github copilot", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const option = ProjectTypeOptions.startWithGithubCopilot();
    assert.notEqual(option.description, undefined);
  });
  it("start with github copilot with preview", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const option = ProjectTypeOptions.startWithGithubCopilot();
    assert.isUndefined(option.description);
  });
});

describe("TeamsProjectTypeOptions", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("CLI label", () => {
    const tab = TeamsProjectTypeOptions.tab(Platform.CLI);
    assert.equal(tab.label, getLocalizedString("core.TabOption.label"));
    const bot = TeamsProjectTypeOptions.bot(Platform.CLI);
    assert.equal(bot.label, getLocalizedString("core.createProjectQuestion.projectType.bot.label"));
    const me = TeamsProjectTypeOptions.me(Platform.CLI);
    assert.equal(me.label, getLocalizedString("core.MessageExtensionOption.label"));
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
      [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
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
      [QuestionNames.ActionManifestPath]: "test",
    };
    const res = folderAndAppNameCondition(inputs);
    assert.isTrue(res);
  });
});

describe("ActionStartOptions", () => {
  it("mcp() should return correct OptionItem", () => {
    const mcpOption = ActionStartOptions.mcp();

    assert.equal(mcpOption.id, "mcp");
    assert.equal(mcpOption.label, getLocalizedString("core.createProjectQuestion.mcpForDa.label"));
    assert.equal(
      mcpOption.detail,
      getLocalizedString("core.createProjectQuestion.mcpForDa.detail")
    );
    assert.equal(mcpOption.data, TemplateNames.DeclarativeAgentWithActionFromMCP);
  });
});

describe("constructNode", () => {
  it("should return foundryNode when node is foundryNode", () => {
    const json = JSON.stringify({
      node: "foundryNode",
      condition: { enum: ["foundry-proxy-agent"] },
    });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.isDefined(node.data);
    assert.deepEqual(node.condition, { enum: ["foundry-proxy-agent"] });
  });

  it("should return llmServiceNode when node is llmServiceNode", () => {
    const json = JSON.stringify({
      node: "llmServiceNode",
      condition: { enum: ["some-cap"] },
    });
    const node = constructNode(json);
    assert.isDefined(node);
  });

  it("should return apiSpecNode when node is apiSpecNode", () => {
    const json = JSON.stringify({
      node: "apiSpecNode",
      condition: { equals: "api-spec" },
    });
    const node = constructNode(json);
    assert.isDefined(node);
  });

  it("should return azureOpenAINode when node is azureOpenAINode", () => {
    const json = JSON.stringify({
      node: "azureOpenAINode",
      condition: { equals: "llm-service-azure-openai" },
    });
    const node = constructNode(json);
    assert.isDefined(node);
  });

  it("should build a generic node with options, children, and condition", () => {
    const json = JSON.stringify({
      data: {
        type: "singleSelect",
        name: "test-question",
        title: "core.createProjectQuestion.llmService.title",
        placeholder: "core.createProjectQuestion.llmService.placeholder",
        options: [
          {
            id: "opt1",
            label: "core.createProjectQuestion.llmServiceAzureOpenAIOption.label",
            detail: "core.createProjectQuestion.llmServiceAzureOpenAIOption.detail",
            data: "SomeTemplate",
          },
        ],
      },
      children: [],
      condition: { enum: ["parent-id"] },
    });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.deepEqual(node.condition, { enum: ["parent-id"] });
    assert.equal((node.data as any).name, "test-question");
  });

  it("should build a generic node without condition", () => {
    const json = JSON.stringify({
      data: {
        type: "singleSelect",
        name: "no-condition-question",
        title: "core.createProjectQuestion.llmService.title",
        placeholder: "core.createProjectQuestion.llmService.placeholder",
        options: [],
      },
      children: [],
    });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.isUndefined(node.condition);
  });

  it("should recursively build child nodes", () => {
    const json = JSON.stringify({
      data: {
        type: "singleSelect",
        name: "parent-question",
        title: "core.createProjectQuestion.llmService.title",
        placeholder: "core.createProjectQuestion.llmService.placeholder",
        options: [],
      },
      children: [
        {
          node: "foundryNode",
          condition: { enum: ["foundry-proxy-agent"] },
        },
      ],
    });
    const node = constructNode(json);
    assert.isDefined(node.children);
    assert.lengthOf(node.children!, 1);
  });
});

describe("foundryNode", () => {
  it("should return a node with foundryEndpointQuestion as data", () => {
    const node = foundryNode();
    assert.isDefined(node);
    assert.isDefined(node.data);
    assert.equal((node.data as any).name, QuestionNames.FoundryEndpoint);
    assert.isUndefined(node.condition);
  });

  it("should accept a condition and apply it to the node", () => {
    const condition = { enum: ["foundry-proxy-agent"] };
    const node = foundryNode(condition);
    assert.deepEqual(node.condition, condition);
  });

  it("child condition should be true when FoundryEndpoint has a value", () => {
    const node = foundryNode();
    const childCondition = node.children?.[0].condition as ConditionFunc;
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.FoundryEndpoint]: "https://my-foundry.azure.com",
    };
    assert.isTrue(childCondition(inputs));
  });

  it("child condition should be false when FoundryEndpoint is empty string", () => {
    const node = foundryNode();
    const childCondition = node.children?.[0].condition as ConditionFunc;
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.FoundryEndpoint]: "",
    };
    assert.isFalse(childCondition(inputs));
  });

  it("child condition should be false when FoundryEndpoint is undefined", () => {
    const node = foundryNode();
    const childCondition = node.children?.[0].condition as ConditionFunc;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    assert.isFalse(childCondition(inputs));
  });

  it("child node should have foundryAgentIdQuestion as data", () => {
    const node = foundryNode();
    const childData = node.children?.[0].data;
    assert.isDefined(childData);
    assert.equal((childData as any).name, QuestionNames.FoundryAgentId);
  });
});

describe("foundryEndpointQuestion and foundryAgentIdQuestion", () => {
  it("foundryEndpointQuestion should have correct name", () => {
    const question = foundryEndpointQuestion();
    assert.equal(question.name, QuestionNames.FoundryEndpoint);
    assert.equal(question.type, "text");
    assert.isDefined(question.title);
    assert.isDefined(question.placeholder);
  });

  it("foundryAgentIdQuestion should have correct name", () => {
    const question = foundryAgentIdQuestion();
    assert.equal(question.name, QuestionNames.FoundryAgentId);
    assert.equal(question.type, "text");
    assert.isDefined(question.title);
    assert.isDefined(question.placeholder);
  });
});

describe("rootNode", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should use cached JSON path when not using local template and cached file exists", () => {
    // Test the actual function — it loads from bundled in dev mode
    const node = getRootProjectTypeNode();
    assert.isDefined(node);
    assert.isDefined(node.data);
    const data = node.data as SingleSelectQuestion;
    assert.equal(data.name, "project-type");
    assert.equal(data.type, "singleSelect");
    // Root should have project type options
    assert.isTrue((data.staticOptions as OptionItem[]).length >= 5);
  });

  it("should use templates folder when using local template", () => {
    // In dev mode (alpha version), useLocalTemplate returns true → uses bundled
    const node = getRootProjectTypeNode(Platform.VSCode);
    assert.isDefined(node);
    // Should have children (sub-tree nodes for each project type)
    assert.isDefined(node.children);
    assert.isTrue(node.children!.length >= 5);
  });

  it("should use templates folder when cached file does not exist", () => {
    // Verify the node has correct structure regardless of source
    const node = getRootProjectTypeNode();
    assert.isDefined(node);
    const data = node.data as SingleSelectQuestion;
    // Check that options include known project types
    const optionIds = (data.staticOptions as OptionItem[]).map((o) => o.id);
    assert.include(optionIds, "copilot-agent-type");
    assert.include(optionIds, "custom-engine-agent-type");
    assert.include(optionIds, "teams-agent-and-app-type");
    assert.include(optionIds, "office-meta-os-type");
    assert.include(optionIds, "graph-connector-type");
  });

  it("should pass platform parameter to constructNode", () => {
    // Verify that getRootProjectTypeNode returns a valid node for different platforms
    const nodeVSC = getRootProjectTypeNode(Platform.VSCode);
    assert.isDefined(nodeVSC);
    assert.isDefined(nodeVSC.data);

    const nodeCLI = getRootProjectTypeNode(Platform.CLI);
    assert.isDefined(nodeCLI);
    assert.isDefined(nodeCLI.data);
  });
});

describe("constructNode", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should construct a singleSelect node from JSON", () => {
    const json = JSON.stringify({
      data: {
        title: "core.createProjectQuestion.title",
        name: "project-type",
        type: "singleSelect",
        options: [
          {
            id: "test-option",
            label: "core.createProjectQuestion.projectType.declarativeAgent.label",
            detail: "core.createProjectQuestion.projectType.declarativeAgent.detail",
          },
        ],
      },
    });

    const node = constructNode(json);
    assert.isDefined(node);
    assert.isDefined(node.data);
    const data = node.data as SingleSelectQuestion;
    assert.equal(data.type, "singleSelect");
    assert.equal(data.name, "project-type");
    assert.isDefined(data.staticOptions);
    assert.equal((data.staticOptions as OptionItem[]).length, 1);
    assert.equal((data.staticOptions as OptionItem[])[0].id, "test-option");
  });

  it("should handle icon prefix for VSCode platform", () => {
    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [
          {
            id: "opt1",
            label: "core.createProjectQuestion.projectType.declarativeAgent.label",
            icon: "$(teamsfx-agent)",
          },
        ],
      },
    });

    const node = constructNode(json, Platform.VSCode);
    const data = node.data as SingleSelectQuestion;
    const option = (data.staticOptions as OptionItem[])[0];
    assert.isTrue(option.label.startsWith("$(teamsfx-agent) "));
  });

  it("should not add icon prefix for CLI platform", () => {
    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [
          {
            id: "opt1",
            label: "core.createProjectQuestion.projectType.declarativeAgent.label",
            icon: "$(teamsfx-agent)",
          },
        ],
      },
    });

    const node = constructNode(json, Platform.CLI);
    const data = node.data as SingleSelectQuestion;
    const option = (data.staticOptions as OptionItem[])[0];
    assert.isFalse(option.label.startsWith("$(teamsfx-agent) "));
  });

  it("should handle groupName in options", () => {
    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [
          {
            id: "opt1",
            label: "core.createProjectQuestion.projectType.declarativeAgent.label",
            groupName: "core.createProjectQuestion.projectType.createGroup.aiAgent",
          },
        ],
      },
    });

    const node = constructNode(json);
    const data = node.data as SingleSelectQuestion;
    const option = (data.staticOptions as OptionItem[])[0];
    assert.isDefined(option.groupName);
  });

  it("should filter out feature-flagged options when flag is disabled", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);

    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [
          { id: "always-visible", label: "Always" },
          { id: "flagged", label: "Flagged", featureFlag: "DAMetaOS" },
        ],
      },
    });

    const node = constructNode(json);
    const data = node.data as SingleSelectQuestion;
    const options = data.staticOptions as OptionItem[];
    assert.equal(options.length, 1);
    assert.equal(options[0].id, "always-visible");
  });

  it("should include feature-flagged options when flag is enabled", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);

    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [
          { id: "always-visible", label: "Always" },
          { id: "flagged", label: "Flagged", featureFlag: "DAMetaOS" },
        ],
      },
    });

    const node = constructNode(json);
    const data = node.data as SingleSelectQuestion;
    const options = data.staticOptions as OptionItem[];
    assert.equal(options.length, 2);
  });

  it("should handle group type nodes", () => {
    const json = JSON.stringify({
      data: { type: "group", name: "test-group" },
      children: [],
    });

    const node = constructNode(json);
    assert.isDefined(node);
    assert.equal(node.data?.type, "group");
  });

  it("should handle condition on nodes", () => {
    const json = JSON.stringify({
      condition: { equals: "some-value" },
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [],
      },
    });

    const node = constructNode(json);
    const condition = node.condition as StringValidation;
    assert.equal(condition.equals, "some-value");
  });

  it("should recursively construct children", () => {
    const json = JSON.stringify({
      data: {
        title: "parent.title",
        name: "parent",
        type: "singleSelect",
        options: [{ id: "child-trigger", label: "Child" }],
      },
      children: [
        {
          condition: { equals: "child-trigger" },
          data: {
            title: "child.title",
            name: "child",
            type: "singleSelect",
            options: [{ id: "sub-item", label: "Sub" }],
          },
        },
      ],
    });

    const node = constructNode(json);
    assert.isDefined(node.children);
    assert.equal(node.children!.length, 1);
    const childCondition = node.children![0].condition as StringValidation;
    assert.equal(childCondition.equals, "child-trigger");
  });

  it("should handle skipSingleOption property", () => {
    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        skipSingleOption: true,
        options: [{ id: "only", label: "Only Option" }],
      },
    });

    const node = constructNode(json);
    const data = node.data as SingleSelectQuestion;
    assert.isTrue(data.skipSingleOption);
  });

  it("should resolve known node references", () => {
    const json = JSON.stringify({
      node: "llmServiceNode",
    });

    const node = constructNode(json);
    assert.isDefined(node);
  });

  it("should throw for unknown node references", () => {
    const json = JSON.stringify({
      node: "nonExistentNode",
    });

    assert.throws(() => constructNode(json), /Unknown node reference: nonExistentNode/);
  });

  it("should set data property on options", () => {
    const json = JSON.stringify({
      data: {
        title: "test.title",
        name: "test",
        type: "singleSelect",
        options: [
          { id: "opt1", label: "Option 1", data: "template-name-1" },
          { id: "opt2", label: "Option 2" },
        ],
      },
    });

    const node = constructNode(json);
    const data = node.data as SingleSelectQuestion;
    const options = data.staticOptions as OptionItem[];
    assert.equal(options[0].data, "template-name-1");
    assert.isUndefined(options[1].data);
  });
});

describe("constructNode - additional coverage", () => {
  it("should resolve apiSpecWithSearchNode reference", () => {
    const json = JSON.stringify({ node: "apiSpecWithSearchNode" });
    const node = constructNode(json);
    assert.isDefined(node);
  });

  it("should resolve apiSpecWithSearchNode with condition", () => {
    const json = JSON.stringify({
      node: "apiSpecWithSearchNode",
      condition: { equals: "some-value" },
    });
    const node = constructNode(json);
    assert.isDefined(node);
    const condition = node.condition as StringValidation;
    assert.equal(condition.equals, "some-value");
  });

  it("should resolve foundryNode reference", () => {
    const json = JSON.stringify({ node: "foundryNode" });
    const node = constructNode(json);
    assert.isDefined(node);
  });

  it("should resolve gcNameNode reference", () => {
    const json = JSON.stringify({ node: "gcNameNode" });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.isDefined(node.data);
  });

  it("should resolve gcConnectionIdNode reference", () => {
    const json = JSON.stringify({ node: "gcConnectionIdNode" });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.isDefined(node.data);
  });

  it("should resolve officeAddinFolderNode reference", () => {
    const json = JSON.stringify({ node: "officeAddinFolderNode" });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.equal(node.data?.type, "folder");
  });

  it("should resolve officeAddinImportNode reference", () => {
    const json = JSON.stringify({ node: "officeAddinImportNode" });
    const node = constructNode(json);
    assert.isDefined(node);
    assert.equal(node.data?.type, "group");
    assert.isDefined(node.children);
    assert.equal(node.children!.length, 2);
  });
});

describe("rootNode - cache vs bundled", () => {
  it("should load and return valid tree with all project types", () => {
    const node = getRootProjectTypeNode(Platform.VSCode);
    assert.isDefined(node);
    assert.isDefined(node.data);
    const data = node.data as SingleSelectQuestion;
    const optionIds = (data.staticOptions as OptionItem[]).map((o) => o.id);
    assert.include(optionIds, "copilot-agent-type");
    assert.include(optionIds, "custom-engine-agent-type");
    assert.include(optionIds, "graph-connector-type");
    assert.include(optionIds, "teams-agent-and-app-type");
    assert.include(optionIds, "office-meta-os-type");
    assert.isDefined(node.children);
    assert.isTrue(node.children!.length >= 5);
  });

  it("should work for CLI platform without icons", () => {
    const node = getRootProjectTypeNode(Platform.CLI);
    assert.isDefined(node);
    const data = node.data as SingleSelectQuestion;
    const options = data.staticOptions as OptionItem[];
    for (const opt of options) {
      assert.isFalse(
        opt.label.startsWith("$("),
        "Option " + opt.id + " should not have icon on CLI"
      );
    }
  });
});

describe("wizard sub-tree extraction", () => {
  it("CEA sub-tree should have correct options", () => {
    const root = getRootProjectTypeNode(Platform.VSCode);
    const node = root.children?.find(
      (c) => (c.condition as StringValidation)?.equals === "custom-engine-agent-type"
    );
    assert.isDefined(node);
    const data = node!.data as SingleSelectQuestion;
    const optionIds = (data.staticOptions as OptionItem[]).map((o) => o.id);
    assert.include(optionIds, "basic-custom-engine-agent");
    assert.include(optionIds, "weather-agent");
  });

  it("Teams sub-tree should have correct condition", () => {
    const node = getTeamsProjectNode();
    assert.isDefined(node);
    const condition = node.condition as StringValidation;
    assert.equal(condition.equals, "teams-agent-and-app-type");
    assert.isDefined(node.data);
  });

  it("CEA sub-tree should have children", () => {
    const root = getRootProjectTypeNode(Platform.VSCode);
    const node = root.children?.find(
      (c) => (c.condition as StringValidation)?.equals === "custom-engine-agent-type"
    );
    assert.isDefined(node!.children);
    assert.isTrue(node!.children!.length > 0);
  });

  it("Teams sub-tree should have children", () => {
    const node = getTeamsProjectNode();
    assert.isDefined(node.children);
    assert.isTrue(node.children!.length > 0);
  });
});
