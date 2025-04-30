// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import {
  ConversationState,
  MemoryStorage,
  TurnContext,
  UserState,
} from "@microsoft/agents-hosting";
import * as sinon from "sinon";
import { CustomStorage } from "../testUtils";
import { DefaultBotSsoExecutionActivityHandler } from "../../../../../src/conversation/sso/defaultBotSsoExecutionActivityHandler";
import { BotSsoExecutionDialog } from "../../../../../src/conversation/sso/botSsoExecutionDialog";
import { BotSsoConfig } from "../../../../../src/conversation/interface";

// eslint-disable-next-line no-secrets/no-secrets
describe("DefaultBotSsoExecutionActivityHandler Tests - Node", () => {
  const sandbox = sinon.createSandbox();

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

  afterEach(() => {
    sandbox.restore();
  });

  it("create default sso execution activity handler should work", () => {
    const defaultBotSsoExecutionActivityHandler: any = new DefaultBotSsoExecutionActivityHandler(
      ssoConfig
    );
    const ssoExecutionDialog = defaultBotSsoExecutionActivityHandler.ssoExecutionDialog;
    assert.isDefined(ssoExecutionDialog);

    const userState = defaultBotSsoExecutionActivityHandler.userState;
    assert.isDefined(userState);
    assert.isTrue(userState.storage instanceof MemoryStorage);

    const conversationState = defaultBotSsoExecutionActivityHandler.conversationState;
    assert.isDefined(conversationState);
    assert.isTrue(userState.storage instanceof MemoryStorage);

    const dialogState = defaultBotSsoExecutionActivityHandler.dialogState;
    assert.isDefined(dialogState);
  });

  it("create default sso execution activity handler should work with custom config", () => {
    const storage = new CustomStorage();

    const ssoConfig: BotSsoConfig = {
      dialog: {
        CustomBotSsoExecutionActivityHandler: DefaultBotSsoExecutionActivityHandler,
        userState: new UserState(storage),
        conversationState: new ConversationState(storage),
        dedupStorage: storage,
      },
      aad: {
        scopes: ["User.Read"],
        clientId,
        clientSecret,
        tenantId,
        authorityHost,
        initiateLoginEndpoint,
      },
    };
    const defaultBotSsoExecutionActivityHandler: any = new DefaultBotSsoExecutionActivityHandler(
      ssoConfig
    );

    const ssoExecutionDialog = defaultBotSsoExecutionActivityHandler.ssoExecutionDialog;
    assert.isDefined(ssoExecutionDialog);

    const userState = defaultBotSsoExecutionActivityHandler.userState;
    assert.isDefined(userState);
    assert.isTrue(userState.storage instanceof CustomStorage);

    const conversationState = defaultBotSsoExecutionActivityHandler.conversationState;
    assert.isDefined(conversationState);
    assert.isTrue(userState.storage instanceof CustomStorage);
  });

  it("trigger sign in function should call sso execution dialog", () => {
    const defaultBotSsoExecutionActivityHandler: any = new DefaultBotSsoExecutionActivityHandler(
      ssoConfig
    );
    const ssoExecutionDialog =
      defaultBotSsoExecutionActivityHandler.ssoExecutionDialog as BotSsoExecutionDialog;
    const stub = sinon.stub(ssoExecutionDialog, "run").resolves();
    const context = sandbox.createStubInstance(TurnContext);

    defaultBotSsoExecutionActivityHandler.handleTeamsSigninVerifyState(context);
    assert.isTrue(stub.callCount === 1);

    defaultBotSsoExecutionActivityHandler.handleTeamsSigninTokenExchange(context);
    assert.isTrue(stub.callCount === 2);
  });
});
