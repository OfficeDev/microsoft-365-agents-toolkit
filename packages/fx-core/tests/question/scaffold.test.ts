// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ConditionFunc,
  FuncValidation,
  Inputs,
  Platform,
  TextInputQuestion,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import { environmentNameManager } from "../../src/core/environmentName";
import { QuestionNames } from "../../src/question/constants";
import {
  addAuthActionQuestion,
  apiFromPluginManifestQuestion,
  apiSpecFromPluginManifestQuestion,
  kiotaRegenerateQuestion,
  selectTargetEnvQuestion,
} from "../../src/question/other";
import * as sinon from "sinon";
import fs from "fs-extra";
import { Bot } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/bot";
import { ConfigurableTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/configurableTab";
import { CommandScope, MeetingsContext } from "../../src/component/driver/teamsApp/utils/utils";
import { StaticTab } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/staticTab";
import { MessagingExtension } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/messagingExtension";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { getTemplateName } from "../../src/question/scaffold/vsc/createFromTdpNode";
import { TemplateNames } from "../../src/question/templates";

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

  const validConfigurableTabForTabCode: ConfigurableTab = {
    objectId: "objId",
    configurationUrl: "https://url",
    canUpdateConfiguration: false,
    scopes: [CommandScope.GroupChat],
    context: [MeetingsContext.ChannelTab],
    sharePointPreviewImage: "img",
    supportedSharePointHosts: [],
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
    chai.assert.equal(res, TemplateNames.TabAndDefaultBot);
  });

  it("return TabNonSso", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      staticTabs: [validStaticTab],
    };

    const res = getTemplateName(appDefinition);
    chai.assert.equal(res, TemplateNames.Tab);
  });

  it("return DefaultBotAndMessageExtension", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      bots: [validBot],
      messagingExtensions: [validMessagingExtension],
    };

    const res = getTemplateName(appDefinition);
    chai.assert.equal(res, TemplateNames.BotAndMessageExtension);
  });

  it("return MessageExtension", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      messagingExtensions: [validMessagingExtension],
    };

    const res = getTemplateName(appDefinition);
    chai.assert.equal(res, TemplateNames.MessageExtension);
  });

  it("return bot", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
      bots: [validBot],
    };

    const res = getTemplateName(appDefinition);
    chai.assert.equal(res, TemplateNames.DefaultBot);
  });

  it("return undefined", () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "id",
    };

    const res = getTemplateName(appDefinition);
    chai.assert.isUndefined(res);
  });
});
