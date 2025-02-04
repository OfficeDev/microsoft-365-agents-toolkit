// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { assert } from "chai";
import "mocha";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { Bot } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/bot";
import { MessagingExtension } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/messagingExtension";
import { StaticTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/staticTab";
import {
  createFromTdpNode,
  getTemplateName,
} from "../../src/question/scaffold/vsc/createFromTdpNode";
import { TemplateNames } from "../../src/question/templates";
import { scaffoldQuestionForVSCode } from "../../src/question/scaffold/vsc/createRootNode";
import { scaffoldQuestionForVS } from "../../src/question/scaffold/vs/createRootNode";
import { apiSpecNode } from "../../src/question/scaffold/vsc/m365ProjectTypeNode";
import { ConditionFunc, Inputs, Platform } from "@microsoft/teamsfx-api";
import { QuestionNames } from "../../src/question/constants";

describe("vsc", () => {
  it("scaffoldQuestionForVSCode", () => {
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

    const res = getTemplateName(appDefinition);
    assert.equal(res, TemplateNames.TabAndDefaultBot);
  });

  it("return TabNonSso", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      staticTabs: [validStaticTab],
    };

    const res = getTemplateName(appDefinition);
    assert.equal(res, TemplateNames.Tab);
  });

  it("return DefaultBotAndMessageExtension", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      bots: [validBot],
      messagingExtensions: [validMessagingExtension],
    };

    const res = getTemplateName(appDefinition);
    assert.equal(res, TemplateNames.BotAndMessageExtension);
  });

  it("return MessageExtension", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      messagingExtensions: [validMessagingExtension],
    };

    const res = getTemplateName(appDefinition);
    assert.equal(res, TemplateNames.MessageExtension);
  });

  it("return bot", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      bots: [validBot],
    };

    const res = getTemplateName(appDefinition);
    assert.equal(res, TemplateNames.DefaultBot);
  });

  it("return undefined", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
    };

    const res = getTemplateName(appDefinition);
    assert.isUndefined(res);
  });
});

describe("m365ProjectTypeNode", () => {
  it("apiSpecNode", () => {
    const node = apiSpecNode({ equals: "a" });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.ApiPluginManifestPath]: "a",
    };
    const condition = node.children?.[1].condition as ConditionFunc;
    const res = condition?.(inputs);
    assert.isTrue(res);
  });
});
