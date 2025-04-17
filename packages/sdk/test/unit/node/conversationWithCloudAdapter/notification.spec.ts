/* eslint-disable @typescript-eslint/no-explicit-any */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CloudAdapter,
  CardFactory,
  TurnContext,
  TurnContextStateCollection,
  ConnectorClient,
} from "@microsoft/agents-hosting";
import {
  TeamDetails,
  TeamsChannelAccount,
  TeamsInfo,
  ChannelInfo,
} from "@microsoft/agents-hosting-teams";
import { Activity, ConversationReference } from "@microsoft/agents-activity";
import { assert, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import * as sinon from "sinon";
import * as fs from "fs";
import * as path from "path";
import { NotificationTargetType } from "../../../../src/conversation/interface";
import { NotificationMiddleware } from "../../../../src/conversation/middlewares/notificationMiddleware";
import {
  Channel,
  Member,
  NotificationBot,
  sendAdaptiveCard,
  sendMessage,
  TeamsBotInstallation,
} from "../../../../src/conversationWithCloudAdapter/notification";
import * as utils from "../../../../src/conversation/utils";
import { TestTarget } from "../conversation/testUtils";
import { DefaultConversationReferenceStore } from "../../../../src/conversation/storage";

chaiUse(chaiPromises);

describe("Notification Tests - Node", () => {
  it("sendMessage should send correct text", async () => {
    const target = new TestTarget();
    await sendMessage(target, "test");
    assert.strictEqual(target.content, "test");
  });

  it("sendMessage should catch and handle error", async () => {
    const target = new TestTarget();
    target.error = new Error("test");
    let errorMessage = "";
    await sendMessage(target, "test", (ctx, err) => {
      errorMessage = err.message;
      return Promise.resolve();
    });
    assert.strictEqual(errorMessage, "test");
  });

  it("sendAdaptiveCard should send correct card", async () => {
    const target = new TestTarget();
    await sendAdaptiveCard(target, { foo: "bar" });
    assert.deepStrictEqual(target.content, { foo: "bar" });
  });

  it("sendAdaptiveCard should catch and handle error", async () => {
    const target = new TestTarget();
    target.error = new Error("test");
    let errorMessage = "";
    await sendAdaptiveCard(target, { foo: "bar" }, (ctx, err) => {
      errorMessage = err.message;
      return Promise.resolve();
    });
    assert.strictEqual(errorMessage, "test");
  });

  describe("Channel Tests - Node", () => {
    const sandbox = sinon.createSandbox();
    const fakeBotAppId = "fakeBotAppId";
    let botInstallation: TeamsBotInstallation;
    let content: any;
    let activityResponse: any;
    let turnError: Error | undefined;

    beforeEach(() => {
      content = "";
      activityResponse = {};
      turnError = undefined;
      const stubContext = sandbox.createStubInstance(TurnContext);
      stubContext.sendActivity.callsFake((activityOrText, speak, inputHint) => {
        if (turnError) {
          throw turnError;
        }
        return new Promise((resolve) => {
          content = activityOrText;
          resolve(activityResponse);
        });
      });
      const stubAdapter = sandbox.createStubInstance(CloudAdapter);
      (
        stubAdapter.continueConversation as unknown as sinon.SinonStub<
          [Partial<ConversationReference>, (context: TurnContext) => Promise<void>],
          Promise<void>
        >
      ).callsFake(async (ref, logic) => {
        await logic(stubContext);
      });
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const stubConnectorClient = sandbox.createStubInstance(ConnectorClient);
      stubConnectorClient.createConversationAsync.callsFake(() => {
        return { id: "1" } as any;
      });
      const stubTurnState = sandbox.createStubInstance(TurnContextStateCollection);
      stubTurnState.get.returns(stubConnectorClient);
      sandbox.stub(TurnContext.prototype, "turnState").get(() => stubTurnState);
      sandbox.stub(TurnContext.prototype, "activity").get(() => {
        return {
          conversation: {
            tenantId: "11",
          },
          recipient: {},
          getConversationReference: () => {
            return {
              conversation: {
                id: "1",
              },
            };
          },
        };
      });
      botInstallation = new TeamsBotInstallation(stubAdapter, conversationRef as any, fakeBotAppId);
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("sendMessage should send correct text", async () => {
      const channel = new Channel(botInstallation, { id: "1" } as ChannelInfo);
      assert.strictEqual(channel.type, "Channel");
      activityResponse = {
        id: "message-x",
      };
      let res = await channel.sendMessage("text");
      assert.strictEqual(content, "text");
      assert.deepStrictEqual(res, { id: "message-x" });
      activityResponse = undefined;
      res = await channel.sendMessage("text");
      assert.deepStrictEqual(res, { id: undefined });
    });

    it("sendMessage should handle error", async () => {
      const channel = new Channel(botInstallation, { id: "1" } as ChannelInfo);
      activityResponse = {
        id: "message-x",
      };
      turnError = new Error("error-message-x");
      let errorMessage = "";
      await channel.sendMessage("text", (ctx, err) => {
        errorMessage = err.message;
        return Promise.resolve();
      });
      assert.strictEqual(errorMessage, "error-message-x");
    });

    it("sendMessage should throw error if onError is undefined", async () => {
      const channel = new Channel(botInstallation, { id: "1" } as ChannelInfo);
      activityResponse = {
        id: "message-x",
      };
      turnError = new Error("error-message-x");
      let actualError: Error | undefined;
      try {
        await channel.sendMessage("text");
      } catch (error) {
        actualError = error as Error;
      }
      assert.isDefined(actualError);
    });

    it("sendAdaptiveCard should send correct card", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const channel = new Channel(botInstallation, { id: "1" } as ChannelInfo);
      assert.strictEqual(channel.type, "Channel");
      activityResponse = {
        id: "message-x",
      };
      let res = await channel.sendAdaptiveCard({ foo: "bar" });
      assert.deepStrictEqual(content, {
        attachments: [
          {
            content: {
              foo: "bar",
            },
          },
        ],
        type: "message",
      });
      assert.deepStrictEqual(res, { id: "message-x" });
      activityResponse = undefined;
      res = await channel.sendAdaptiveCard({ foo: "bar" });
      assert.deepStrictEqual(res, { id: undefined });
    });

    it("sendAdaptiveCard should handle error", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const channel = new Channel(botInstallation, { id: "1" } as ChannelInfo);
      activityResponse = {
        id: "message-x",
      };
      turnError = new Error("error-card-x");
      let errorMessage = "";
      await channel.sendAdaptiveCard({ foo: "bar" }, (ctx, err) => {
        errorMessage = err.message;
        return Promise.resolve();
      });
      assert.strictEqual(errorMessage, "error-card-x");
    });

    it("sendAdaptiveCard should throw error if onError is undefined", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const channel = new Channel(botInstallation, { id: "1" } as ChannelInfo);
      activityResponse = {
        id: "message-x",
      };
      turnError = new Error("error-card-x");
      let actualError: Error | undefined;
      try {
        await channel.sendAdaptiveCard({ foo: "bar" });
      } catch (error) {
        actualError = error as Error;
      }
      assert.isDefined(actualError);
    });
  });

  describe("Member Tests - Node", () => {
    const sandbox = sinon.createSandbox();
    let botInstallation: TeamsBotInstallation;
    let content: any;
    let activityResponse: any;
    let turnError: Error | undefined;

    beforeEach(() => {
      content = "";
      activityResponse = {};
      turnError = undefined;
      const fakeBotAppId = "fakeBotAppId";
      const stubConnectorClient = sandbox.createStubInstance(ConnectorClient);
      stubConnectorClient.createConversationAsync.callsFake(() => {
        return { id: "1" } as any;
      });
      const stubTurnState = sandbox.createStubInstance(TurnContextStateCollection);
      stubTurnState.get.returns(stubConnectorClient);
      const stubContext = sandbox.createStubInstance(TurnContext);
      stubContext.sendActivity.callsFake((activityOrText, speak, inputHint) => {
        if (turnError) {
          throw turnError;
        }
        return new Promise((resolve) => {
          content = activityOrText;
          resolve(activityResponse);
        });
      });
      sandbox.stub(TurnContext.prototype, "turnState").get(() => stubTurnState);
      sandbox.stub(TurnContext.prototype, "activity").get(() => {
        return {
          conversation: {
            tenantId: "11",
          },
          recipient: {},
          getConversationReference: () => {
            return {
              conversation: {
                id: "1",
              },
            };
          },
        };
      });
      const stubAdapter = sandbox.createStubInstance(CloudAdapter);
      (
        stubAdapter.continueConversation as unknown as sinon.SinonStub<
          [Partial<ConversationReference>, (context: TurnContext) => Promise<void>],
          Promise<void>
        >
      ).callsFake(async (ref, logic) => {
        await logic(stubContext);
      });
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      botInstallation = new TeamsBotInstallation(stubAdapter, conversationRef as any, fakeBotAppId);
      sandbox
        .stub(ConnectorClient.prototype, "createConversationAsync")
        .resolves({ id: "1" } as any);
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("sendMessage should send correct text", async () => {
      const member = new Member(botInstallation, { id: "1" } as TeamsChannelAccount);
      assert.strictEqual(member.type, "Person");
      activityResponse = {
        id: "message-y",
      };
      let res = await member.sendMessage("text");
      assert.strictEqual(content, "text");
      assert.deepStrictEqual(res, { id: "message-y" });
      activityResponse = undefined;
      res = await member.sendMessage("text");
      assert.deepStrictEqual(res, { id: undefined });
    });

    it("sendMessage should handle error", async () => {
      const member = new Member(botInstallation, { id: "1" } as TeamsChannelAccount);
      assert.strictEqual(member.type, "Person");
      activityResponse = {
        id: "message-y",
      };
      turnError = new Error("error-message-y");
      let errorMessage = "";
      await member.sendMessage("text", (ctx, err) => {
        errorMessage = err.message;
        return Promise.resolve();
      });
      assert.strictEqual(errorMessage, "error-message-y");
    });

    it("sendMessage should throw error if onError is undefined", async () => {
      const member = new Member(botInstallation, { id: "1" } as TeamsChannelAccount);
      assert.strictEqual(member.type, "Person");
      activityResponse = {
        id: "message-y",
      };
      turnError = new Error("error-message-y");
      let actualError: Error | undefined;
      try {
        await member.sendMessage("text");
      } catch (error) {
        actualError = error as Error;
      }
      assert.isDefined(actualError);
    });

    it("sendAdaptiveCard should send correct card", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const member = new Member(botInstallation, { id: "1" } as TeamsChannelAccount);
      assert.strictEqual(member.type, "Person");
      activityResponse = {
        id: "message-y",
      };
      let res = await member.sendAdaptiveCard({ foo: "bar" });
      assert.deepStrictEqual(content, {
        attachments: [
          {
            content: {
              foo: "bar",
            },
          },
        ],
        type: "message",
      });
      assert.deepStrictEqual(res, { id: "message-y" });
      activityResponse = undefined;
      res = await member.sendAdaptiveCard({ foo: "bar" });
      assert.deepStrictEqual(res, { id: undefined });
    });

    it("sendAdaptiveCard should handle error", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const member = new Member(botInstallation, { id: "1" } as TeamsChannelAccount);
      activityResponse = {
        id: "message-y",
      };
      turnError = new Error("error-card-y");
      let errorMessage = "";
      await member.sendAdaptiveCard({ foo: "bar" }, (ctx, err) => {
        errorMessage = err.message;
        return Promise.resolve();
      });
      assert.strictEqual(errorMessage, "error-card-y");
    });

    it("sendAdaptiveCard should throw error if onError is undefined", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const member = new Member(botInstallation, { id: "1" } as TeamsChannelAccount);
      activityResponse = {
        id: "message-y",
      };
      turnError = new Error("error-card-x");
      let actualError: Error | undefined;
      try {
        await member.sendAdaptiveCard({ foo: "bar" });
      } catch (error) {
        actualError = error as Error;
      }
      assert.isDefined(actualError);
    });
  });

  describe("TeamsBotInstallation Tests - Node", () => {
    const sandbox = sinon.createSandbox();
    let adapter: CloudAdapter;
    let context: TurnContext;
    let content: any;
    let activityResponse: any;
    let turnError: Error | undefined;

    beforeEach(() => {
      content = "";
      activityResponse = {};
      turnError = undefined;
      const stubAdapter = sandbox.createStubInstance(CloudAdapter);
      (
        stubAdapter.continueConversation as unknown as sinon.SinonStub<
          [Partial<ConversationReference>, (context: TurnContext) => Promise<void>],
          Promise<void>
        >
      ).callsFake(async (ref, logic) => {
        await logic(context);
      });
      adapter = stubAdapter;
      const stubContext = sandbox.createStubInstance(TurnContext);
      stubContext.sendActivity.callsFake((activityOrText, speak, inputHint) => {
        if (turnError) {
          throw turnError;
        }
        return new Promise((resolve) => {
          content = activityOrText;
          resolve(activityResponse);
        });
      });
      const stubConnectorClient = sandbox.createStubInstance(ConnectorClient);
      stubConnectorClient.createConversationAsync.callsFake(() => {
        return { id: "1" } as any;
      });
      const stubTurnState = sandbox.createStubInstance(TurnContextStateCollection);
      stubTurnState.get.returns(stubConnectorClient);
      sandbox.stub(TurnContext.prototype, "turnState").get(() => stubTurnState);
      sandbox.stub(TurnContext.prototype, "activity").get(() => {
        return {
          conversation: {
            tenantId: "11",
          },
          recipient: {},
          getConversationReference: () => {
            return {
              conversation: {
                id: "1",
              },
            };
          },
        };
      });
      context = stubContext;
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("sendMessage should send correct text", async () => {
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      assert.strictEqual(installation.type, NotificationTargetType.Channel);
      assert.isTrue(installation.type === "Channel");
      activityResponse = {
        id: "message-a",
      };
      let res = await installation.sendMessage("text");
      assert.strictEqual(content, "text");
      assert.deepStrictEqual(res, { id: "message-a" });
      activityResponse = undefined;
      res = await installation.sendMessage("text");
      assert.deepStrictEqual(res, { id: undefined });
    });

    it("sendMessage should handle error", async () => {
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      activityResponse = {
        id: "message-a",
      };
      turnError = new Error("error-message-a");
      let errorMessage = "";
      await installation.sendMessage("text", (ctx, err) => {
        errorMessage = err.message;
        return Promise.resolve();
      });
      assert.strictEqual(errorMessage, "error-message-a");
    });

    it("sendMessage should throw error if onError is undefined", async () => {
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      activityResponse = {
        id: "message-a",
      };
      turnError = new Error("error-message-a");
      let actualError: Error | undefined;
      try {
        await installation.sendMessage("text");
      } catch (error) {
        actualError = error as Error;
      }
      assert.isDefined(actualError);
    });

    it("sendAdaptiveCard should send correct card", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      assert.strictEqual(installation.type, NotificationTargetType.Channel);
      assert.isTrue(installation.type === "Channel");
      activityResponse = {
        id: "message-a",
      };
      let res = await installation.sendAdaptiveCard({ foo: "bar" });
      assert.deepStrictEqual(content, {
        attachments: [
          {
            content: {
              foo: "bar",
            },
          },
        ],
        type: "message",
      });
      assert.deepStrictEqual(res, { id: "message-a" });
      activityResponse = undefined;
      res = await installation.sendAdaptiveCard({ foo: "bar" });
      assert.deepStrictEqual(res, { id: undefined });
    });

    it("sendAdaptiveCard should handle error", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      activityResponse = {
        id: "message-a",
      };
      turnError = new Error("error-card-a");
      let errorMessage = "";
      await installation.sendAdaptiveCard({ foo: "bar" }, (ctx, err) => {
        errorMessage = err.message;
        return Promise.resolve();
      });
      assert.strictEqual(errorMessage, "error-card-a");
    });

    it("sendAdaptiveCard should throw error if onError is undefined", async () => {
      sandbox.stub(CardFactory, "adaptiveCard").callsFake((card) => {
        return { content: card } as any;
      });
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      activityResponse = {
        id: "message-a",
      };
      turnError = new Error("error-card-a");
      let actualError: Error | undefined;
      try {
        await installation.sendAdaptiveCard({ foo: "bar" });
      } catch (error) {
        actualError = error as Error;
      }
      assert.isDefined(actualError);
    });

    it("channels should return correct channels", async () => {
      sandbox.stub(utils, "getTeamsBotInstallationId").returns("test");
      sandbox.stub(TeamsInfo, "getTeamChannels").resolves([{} as ChannelInfo, {} as ChannelInfo]);
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      assert.strictEqual(installation.type, NotificationTargetType.Channel);
      assert.isTrue(installation.type === "Channel");
      const channels = await installation.channels();
      assert.strictEqual(channels.length, 2);
    });

    it("channels should return empty array if no teamId", async () => {
      sandbox.stub(utils, "getTeamsBotInstallationId").returns(undefined);
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      assert.strictEqual(installation.type, NotificationTargetType.Channel);
      assert.isTrue(installation.type === "Channel");
      const channels = await installation.channels();
      assert.strictEqual(channels.length, 0);
    });

    it("channels should return empty array if conversation type is not channel", async () => {
      sandbox.stub(utils, "getTeamsBotInstallationId").returns("test");
      sandbox.stub(TeamsInfo, "getTeamChannels").resolves([{} as ChannelInfo, {} as ChannelInfo]);

      const conversationRef = {
        conversation: {
          conversationType: "personal",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      assert.isTrue(installation.type !== "Channel");
      const channels = await installation.channels();
      assert.strictEqual(channels.length, 0);
    });

    it("getPagedMembers should return correct members", async () => {
      sandbox.stub(TeamsInfo, "getPagedMembers").resolves({
        continuationToken: "token",
        members: [{} as TeamsChannelAccount, {} as TeamsChannelAccount],
      });
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      assert.strictEqual(installation.type, NotificationTargetType.Channel);
      assert.isTrue(installation.type === "Channel");
      const { data: members, continuationToken } = await installation.getPagedMembers();
      assert.strictEqual(members.length, 2);
      assert.strictEqual(continuationToken, "token");
    });

    it("getTeamDetails should return correct team details", async () => {
      sandbox.stub(utils, "getTeamsBotInstallationId").returns("test");
      sandbox
        .stub(TeamsInfo, "getTeamDetails")
        .resolves({ id: "test", name: "test-team" } as TeamDetails);
      const conversationRef = {
        conversation: {
          conversationType: "channel",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      const teamDetails = await installation.getTeamDetails();
      assert.strictEqual(teamDetails?.id, "test");
      assert.strictEqual(teamDetails?.name, "test-team");
    });

    it("getTeamDetails should return undefined if conversation type is not channel", async () => {
      const conversationRef = {
        conversation: {
          conversationType: "personal",
        },
      };
      const fakeBotAppId = "fakeBotAppId";
      const installation = new TeamsBotInstallation(adapter, conversationRef as any, fakeBotAppId);
      const teamDetails = await installation.getTeamDetails();
      assert.isUndefined(teamDetails);
    });
  });
});

describe("Notification Bot Tests - Node", () => {
  const sandbox = sinon.createSandbox();
  const fileDir = "./test/";
  const localFileName = ".notification.localstore.json";
  const filePath = path.join(fileDir, localFileName);
  const rawData = {
    _a_1: {
      channelId: "1",
      conversation: {
        conversationType: "channel",
        id: "1",
        tenantId: "a",
      },
    },
  };
  let adapter: CloudAdapter;
  let store: DefaultConversationReferenceStore;
  let middlewares: any[];

  beforeEach(() => {
    middlewares = [];
    const stubContext = sandbox.createStubInstance(TurnContext);
    const stubAdapter = sandbox.createStubInstance(CloudAdapter);
    stubAdapter.use.callsFake((args) => {
      middlewares.push(args);
      return stubAdapter;
    });
    (
      stubAdapter.continueConversation as unknown as sinon.SinonStub<
        [Partial<ConversationReference>, (context: TurnContext) => Promise<void>],
        Promise<void>
      >
    ).callsFake(async (ref, logic) => {
      await logic(stubContext);
    });
    adapter = stubAdapter;
    store = new DefaultConversationReferenceStore(fileDir);
  });

  afterEach(() => {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
    sandbox.restore();
  });

  it("initialize notification should create correct middleware", () => {
    new NotificationBot(adapter, { store });
    assert.strictEqual(middlewares.length, 1);
    assert.isTrue(middlewares[0] instanceof NotificationMiddleware);
  });

  it("getPagedInstallations should return correct targets", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) => resolve({ continuationToken: "", members: [] }));
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(
      filePath,
      JSON.stringify({
        _a_1: {
          channelId: "1",
          conversation: {
            id: "1",
            tenantId: "a",
          },
        },
        _a_2: {
          channelId: "2",
          conversation: {
            id: "2",
            tenantId: "a",
          },
        },
      })
    );
    const { data: installations, continuationToken } =
      await notificationBot.getPagedInstallations();
    assert.strictEqual(installations.length, 2);
    assert.strictEqual(installations[0].conversationReference.conversation?.id, "1");
    assert.strictEqual(installations[1].conversationReference.conversation?.id, "2");
    assert.strictEqual(continuationToken, "");
  });

  it("getPagedInstallations should remove invalid target", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      throw {
        name: "test",
        message: "test",
        code: "BotNotInConversationRoster",
      };
    });

    const notificationBot = new NotificationBot(adapter, { store });

    fs.writeFileSync(filePath, JSON.stringify(rawData));
    const { data: installations } = await notificationBot.getPagedInstallations();
    const removedData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.strictEqual(installations.length, 0);
    assert.deepStrictEqual(removedData, {});
  });

  it("getPagedInstallations should skip validation", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      throw {
        name: "test",
        message: "test",
        code: "BotNotInConversationRoster",
      };
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(filePath, JSON.stringify(rawData));

    const { data: installations } = await notificationBot.getPagedInstallations(
      undefined,
      undefined,
      false
    );
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.strictEqual(installations.length, 1);
    assert.strictEqual(installations[0].conversationReference.conversation?.id, "1");
    assert.deepStrictEqual(fileData, rawData);
  });

  it("getPagedInstallations should keep valid target", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      throw {
        name: "test",
        message: "test",
        code: "Throttled",
      };
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(filePath, JSON.stringify(rawData));

    const { data: installations } = await notificationBot.getPagedInstallations();
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));

    assert.strictEqual(installations.length, 1);
    assert.strictEqual(installations[0].conversationReference.conversation?.id, "1");
    assert.deepStrictEqual(fileData, rawData);
  });

  it("findMember should return correct member", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) =>
        resolve({
          continuationToken: undefined as unknown as string,
          members: [
            {
              id: "foo",
              name: "foo",
            } as TeamsChannelAccount,
          ],
        })
      );
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(filePath, JSON.stringify(rawData));

    const member = await notificationBot.findMember((m) =>
      Promise.resolve(m.account.name === "foo")
    );
    assert.strictEqual(member?.account.id, "foo");
  });

  it("findMember should return undefined", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) =>
        resolve({
          continuationToken: undefined as unknown as string,
          members: [],
        })
      );
    });

    const notificationBot = new NotificationBot(adapter, { store });
    const member = await notificationBot.findMember((m) =>
      Promise.resolve(m.account.name === "NotFound")
    );
    assert.isUndefined(member);
  });

  it("findAllMembers should return an array of correct members", async () => {
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) =>
        resolve({
          continuationToken: undefined as unknown as string,
          members: [
            {
              email: "a@contoso.com",
            } as TeamsChannelAccount,
            {
              email: "b@contoso.com",
            } as TeamsChannelAccount,
            {
              email: "c@foo.com",
            } as TeamsChannelAccount,
          ],
        })
      );
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(
      filePath,
      JSON.stringify({
        _a_1: {
          channelId: "1",
          conversation: {
            conversationType: "channel",
            id: "1",
            tenantId: "a",
          },
        },
      })
    );

    const members = await notificationBot.findAllMembers((m) =>
      Promise.resolve(m.account.email?.endsWith("contoso.com") === true)
    );
    assert.lengthOf(members, 2);
  });

  it("findChannel should return correct channel", async () => {
    sandbox.stub(utils, "getTeamsBotInstallationId").returns("test");
    sandbox.stub(TeamsInfo, "getTeamDetails").resolves({ id: "test" } as TeamDetails);
    sandbox.stub(TeamsInfo, "getTeamChannels").resolves([{ id: "1" } as ChannelInfo]);
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) =>
        resolve({
          continuationToken: undefined as unknown as string,
          members: [],
        })
      );
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(filePath, JSON.stringify(rawData));

    const channel = await notificationBot.findChannel((c) => Promise.resolve(c.info.id === "1"));
    assert.strictEqual(channel?.info.id, "1");
  });

  it("findAllChannels should return an array of correct channel", async () => {
    sandbox.stub(utils, "getTeamsBotInstallationId").returns("test");
    sandbox.stub(TeamsInfo, "getTeamDetails").resolves({ id: "test" } as TeamDetails);
    sandbox.stub(TeamsInfo, "getTeamChannels").resolves([{ id: "1" } as ChannelInfo]);
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) =>
        resolve({
          continuationToken: undefined as unknown as string,
          members: [],
        })
      );
    });

    const notificationBot = new NotificationBot(adapter, { store });
    fs.writeFileSync(filePath, JSON.stringify(rawData));

    const channels = await notificationBot.findAllChannels((channel, team) =>
      Promise.resolve(team?.id === "test")
    );
    assert.lengthOf(channels, 1);
  });

  it("findChannel should return undefined", async () => {
    sandbox.stub(utils, "getTeamsBotInstallationId").returns("test");
    sandbox.stub(TeamsInfo, "getTeamDetails").resolves({ id: "test" } as TeamDetails);
    sandbox.stub(TeamsInfo, "getTeamChannels").resolves([{ id: "1" } as ChannelInfo]);
    sandbox.stub(TeamsInfo, "getPagedMembers").callsFake((ctx, pageSize, continuationToken) => {
      return new Promise((resolve) =>
        resolve({
          continuationToken: undefined as unknown as string,
          members: [],
        })
      );
    });

    const notificationBot = new NotificationBot(adapter, { store });
    const channel = await notificationBot.findChannel((c) =>
      Promise.resolve(c.info.id === "NotFound")
    );
    assert.isUndefined(channel);
  });

  it("buildTeamsBotInstallation should return correct data", async () => {
    const reference = {
      channelId: "1",
      conversation: {
        conversationType: "channel",
        id: "1",
        tenantId: "a",
      },
    } as ConversationReference;
    const notificationBot = new NotificationBot(adapter, { store });
    const installation = notificationBot.buildTeamsBotInstallation(reference);
    assert.isNotNull(installation);
  });
});
