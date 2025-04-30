// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, expect } from "chai";
import { CloudAdapter, TurnContext } from "@microsoft/agents-hosting";
import { ConversationReference } from "@microsoft/agents-activity";
import * as sinon from "sinon";
import { CommandBot } from "../../../../src/conversationWithCloudAdapter/command";
import { CommandResponseMiddleware } from "../../../../src/conversation/middlewares/commandMiddleware";
import { TestCommandHandler, TestSsoCommandHandler } from "../conversation/testUtils";
import { DefaultBotSsoExecutionActivityHandler } from "../../../../src/conversation/sso/defaultBotSsoExecutionActivityHandler";
import { BotSsoConfig } from "../../../../src";

describe("CommandBot Tests - Node", () => {
  const sandbox = sinon.createSandbox();
  let adapter: CloudAdapter;
  let middlewares: unknown[];

  const clientId = "fake_client_id";
  const clientSecret = "fake_client_secret";
  const tenantId = "fake_tenant";
  const authorityHost = "fake_authority_host";
  const initiateLoginEndpoint = "fake_initiate_login_endpoint";
  const ssoConfig: BotSsoConfig = {
    aad: {
      scopes: ["User.Read"],
      clientId,
      clientSecret,
      tenantId,
      authorityHost,
      initiateLoginEndpoint,
    },
  };

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
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("create command bot should add correct middleware", () => {
    new CommandBot(adapter);
    assert.isTrue(middlewares[0] instanceof CommandResponseMiddleware);
  });

  it("command should be added through registerCommand API", () => {
    const commandBot = new CommandBot(adapter);
    commandBot.registerCommand(new TestCommandHandler("test"));

    assert.isTrue(middlewares[0] instanceof CommandResponseMiddleware);
    const middleware = middlewares[0] as CommandResponseMiddleware;

    assert.isNotEmpty(middleware.commandHandlers);
    assert.isTrue(middleware.commandHandlers.length === 1);
    assert.isTrue(middleware.commandHandlers[0] instanceof TestCommandHandler);
  });

  it("commands should be added through registerCommands API", () => {
    const commandBot = new CommandBot(adapter);
    const stringPattern = "test";
    const regExpPattern = /^test (.*?)$/i;
    commandBot.registerCommands([
      new TestCommandHandler(stringPattern),
      new TestCommandHandler(regExpPattern),
    ]);

    assert.isTrue(middlewares[0] instanceof CommandResponseMiddleware);
    const middleware = middlewares[0] as CommandResponseMiddleware;

    assert.isNotEmpty(middleware.commandHandlers);
    assert.isTrue(middleware.commandHandlers.length === 2);
    assert.isTrue(typeof middleware.commandHandlers[0].triggerPatterns === "string");
    assert.isTrue(middleware.commandHandlers[1].triggerPatterns instanceof RegExp);
  });

  it("create sso command bot should add correct activity handler", () => {
    const defaultSsoHandler = new DefaultBotSsoExecutionActivityHandler(ssoConfig);
    new CommandBot(
      adapter,
      {
        ssoCommands: [new TestSsoCommandHandler("test")],
      },
      defaultSsoHandler
    );
    assert.isTrue(middlewares[0] instanceof CommandResponseMiddleware);
    const middleware = middlewares[0] as CommandResponseMiddleware;

    assert.isDefined(middleware.ssoActivityHandler);
    assert.isTrue(middleware.ssoActivityHandler instanceof DefaultBotSsoExecutionActivityHandler);
    assert.isTrue(middleware.ssoCommandHandlers.length == 1);
  });

  it("add sso command should throw error if sso activity handler is undefined", () => {
    const commandBot = new CommandBot(adapter);
    assert.isUndefined((middlewares[0] as CommandResponseMiddleware).ssoActivityHandler);

    expect(() => {
      commandBot.registerSsoCommand(new TestSsoCommandHandler("test"));
    }).to.throw("Sso command can only be used or added when sso activity handler is not undefined");
  });

  it("add sso command handler should add correct activity handler", () => {
    const commandBot = new CommandBot(
      adapter,
      undefined,
      new DefaultBotSsoExecutionActivityHandler(ssoConfig)
    );
    const middleware = middlewares[0] as CommandResponseMiddleware;
    commandBot.registerSsoCommand(new TestSsoCommandHandler("test"));
    assert.isDefined(middleware.ssoActivityHandler);
    assert.isTrue(middleware.ssoActivityHandler instanceof DefaultBotSsoExecutionActivityHandler);
    assert.isTrue(middleware.ssoCommandHandlers.length == 1);
  });

  it("add sso command handlers should add correct activity handler", () => {
    const commandBot = new CommandBot(
      adapter,
      undefined,
      new DefaultBotSsoExecutionActivityHandler(ssoConfig)
    );
    const middleware = middlewares[0] as CommandResponseMiddleware;
    commandBot.registerSsoCommands([
      new TestSsoCommandHandler("test"),
      new TestSsoCommandHandler("test2"),
    ]);
    assert.isDefined(middleware.ssoActivityHandler);
    assert.isTrue(middleware.ssoActivityHandler instanceof DefaultBotSsoExecutionActivityHandler);
    assert.isTrue(middleware.ssoCommandHandlers.length == 2);
  });

  it("add both normal command and sso command should add correct activity handler", () => {
    const commandBot = new CommandBot(
      adapter,
      undefined,
      new DefaultBotSsoExecutionActivityHandler(ssoConfig)
    );
    const middleware = middlewares[0] as CommandResponseMiddleware;
    commandBot.registerCommand(new TestCommandHandler("test"));
    commandBot.registerSsoCommand(new TestSsoCommandHandler("test"));
    assert.isDefined(middleware.ssoActivityHandler);
    assert.isTrue(middleware.ssoActivityHandler instanceof DefaultBotSsoExecutionActivityHandler);
    assert.isTrue(middleware.commandHandlers.length == 1);
    assert.isTrue(middleware.ssoCommandHandlers.length == 1);
  });
});
