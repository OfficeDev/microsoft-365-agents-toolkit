// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CardFactory,
  ConversationState,
  MemoryStorage,
  AgentStatePropertyAccessor,
  StatusCodes,
  TurnContext,
} from "@microsoft/agents-hosting";
import {
  TeamsChannelAccount,
  TeamsInfo,
  tokenExchangeOperationName,
  verifyStateOperationName,
} from "@microsoft/agents-hosting-teams";
import {
  ActionTypes,
  Activity,
  ActivityTypes,
  Channels,
  InputHints,
} from "@microsoft/agents-activity";
import { DialogSet, DialogState } from "@microsoft/agents-hosting-dialogs";
import {
  OnBehalfOfUserCredential,
  ErrorWithCode,
  ErrorCode,
  TeamsBotSsoPromptSettings,
  BotSsoExecutionDialog,
  TeamsBotSsoPromptTokenResponse,
  CommandMessage,
  TriggerPatterns,
} from "../../../../../src";
import { assert, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import * as sinon from "sinon";
import mockedEnv from "mocked-env";
import { AccessToken } from "@azure/identity";
import { promisify } from "util";
import { TestSsoCommandHandler, TestAdapter } from "../testUtils";
import { ErrorMessage } from "../../../../../src/core/errors";

chaiUse(chaiPromises);
let mockedEnvRestore: () => void;

describe("BotSsoExecutionDialog Tests - Node", () => {
  const sleep = promisify(setTimeout);

  const clientId = "fake_client_id";
  const clientSecret = "fake_client_secret";
  const tenantId = "fake_tenant";
  const userPrincipalName = "fake_userPrincipalName";
  const authorityHost = "fake_authority_host";
  const initiateLoginEndpoint = "fake_initiate_login_endpoint";

  const requiredScopes: string[] = ["User.Read"];
  const expiresOnTimestamp = 12345678;
  const invokeResponseActivityType = "invokeResponse";
  const id = "fake_id";
  const exchangeToken = "fake_exchange_token";

  const OnBehalfOfCredentialAuthConfig = {
    authorityHost: authorityHost,
    clientId: clientId,
    clientSecret: clientSecret,
    tenantId: tenantId,
  };

  /**
   * {
   * "aud": "test_audience",
   * "iss": "https://login.microsoftonline.com/test_aad_id/v2.0",
   * "iat": 1537231048,
   * "nbf": 1537231048,
   * "exp": 1537234948,
   * "aio": "test_aio",
   * "name": "Teams App Framework SDK Unit Test",
   * "oid": "11111111-2222-3333-4444-555555555555",
   * "preferred_username": "test@microsoft.com",
   * "rh": "test_rh",
   * "scp": "access_as_user",
   * "sub": "test_sub",
   * "tid": "test_tenant_id",
   * "uti": "test_uti",
   * "ver": "2.0"
   * }
   */
  const ssoToken =
    // eslint-disable-next-line no-secrets/no-secrets
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOiJ0ZXN0X2F1ZGllbmNlIiwiaXNzIjoiaHR0cHM6Ly9sb2dpbi5taWNyb3NvZnRvbmxpbmUuY29tL3Rlc3RfYWFkX2lkL3YyLjAiLCJpYXQiOjE1MzcyMzEwNDgsIm5iZiI6MTUzNzIzMTA0OCwiZXhwIjoxNTM3MjM0OTQ4LCJhaW8iOiJ0ZXN0X2FpbyIsIm5hbWUiOiJNT0RTIFRvb2xraXQgU0RLIFVuaXQgVGVzdCIsIm9pZCI6IjExMTExMTExLTIyMjItMzMzMy00NDQ0LTU1NTU1NTU1NTU1NSIsInByZWZlcnJlZF91c2VybmFtZSI6InRlc3RAbWljcm9zb2Z0LmNvbSIsInJoIjoidGVzdF9yaCIsInNjcCI6ImFjY2Vzc19hc191c2VyIiwic3ViIjoidGVzdF9zdWIiLCJ0aWQiOiJ0ZXN0X3RlbmFudF9pZCIsInV0aSI6InRlc3RfdXRpIiwidmVyIjoiMi4wIn0.SshbL1xuE1aNZD5swrWOQYgTR9QCNXkZqUebautBvKM";
  const timeoutValue = 50;
  const sleepTimeOffset: number = timeoutValue + 20;
  const sandbox = sinon.createSandbox();
  const testSsoHandlerResponseMessage = "This message is come from sso command dialog";

  beforeEach(function () {
    mockedEnvRestore = mockedEnv({
      INITIATE_LOGIN_ENDPOINT: initiateLoginEndpoint,
      M365_CLIENT_ID: clientId,
      M365_CLIENT_SECRET: clientSecret,
      M365_TENANT_ID: tenantId,
      M365_AUTHORITY_HOST: authorityHost,
    });

    // Mock onBehalfOfUserCredential implementation
    const onBehalfOfUserCredentialStub_GetToken = sandbox.stub(
      OnBehalfOfUserCredential.prototype,
      "getToken"
    );
    onBehalfOfUserCredentialStub_GetToken.onCall(0).callsFake(async () => {
      throw new ErrorWithCode(
        "The user or administrator has not consented to use the application\nFail to get access token because user has not consent scope.",
        ErrorCode.UiRequiredError
      );
    });
    onBehalfOfUserCredentialStub_GetToken.onCall(1).callsFake(async () => {
      return new Promise<AccessToken>((resolve) => {
        resolve({
          token: exchangeToken,
          expiresOnTimestamp: expiresOnTimestamp,
        });
      });
    });

    sandbox.stub(TeamsInfo, "getMember").callsFake(async () => {
      const account: TeamsChannelAccount = {
        id: "fake_id",
        name: "fake_name",
        userPrincipalName: userPrincipalName,
      };
      return account;
    });
  });

  afterEach(function () {
    sandbox.restore();
    mockedEnvRestore();
  });

  it("should add TeamsBotSsoPrompt and WaterfallDialog to the dialog set", () => {
    // Arrange
    const mockStorage = new MemoryStorage();
    const mockSsoPromptSettings: TeamsBotSsoPromptSettings = {
      scopes: requiredScopes,
    };
    const mockDialogName = "testDialog";

    const botSsoExecutionDialog = new BotSsoExecutionDialog(
      mockStorage,
      mockSsoPromptSettings,
      OnBehalfOfCredentialAuthConfig,
      initiateLoginEndpoint,
      mockDialogName
    );

    // Act
    const dialogs = botSsoExecutionDialog.dialogs;

    // Assert
    assert.isNotNull(dialogs.find(mockDialogName + "TeamsFxSsoPrompt"));
    assert.isNotNull(dialogs.find(mockDialogName + "CommandRouteDialog"));
  });

  it("sso execution dialog should response 'Cannot find command' error when command doesn't exist", async function () {
    this.timeout(500);
    const adapter: TestAdapter = await initializeTestEnv();
    const invalidCommand = "InvalidCommand";
    await adapter.send(invalidCommand).assertReply((activity) => {
      assert.strictEqual(activity.type, ActivityTypes.Message);
      assert.strictEqual(activity.text, `Cannot find command: ${invalidCommand}`);
    });
  });

  it("sso execution dialog should end on invalid message when endOnInvalidMessage default to true", async function () {
    this.timeout(500);
    const adapter: TestAdapter = await initializeTestEnv(undefined, true);

    await adapter
      .send("TestCommand")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock User send invalid message
        const messageActivity = createReply(ActivityTypes.Message, activity);
        messageActivity.text = "user sends invalid message during auth flow";
        adapter.send(messageActivity);
      })
      .assertReply((activity) => {
        assert.strictEqual(activity.type, ActivityTypes.Message);
        assert.strictEqual(activity.text, ErrorMessage.FailedToRetrieveSsoToken);
      });
  });

  it("sso execution dialog should timeout with teams verification invoke activity when wait a long time", async function () {
    const adapter: TestAdapter = await initializeTestEnv(timeoutValue);

    await adapter
      .send("TestCommand")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply(async (activity) => {
        // User has not consent. Assert bot send out 412
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.PRECONDITION_FAILED);
        assert.strictEqual(
          (activity.value as any).body.failureDetail,
          "The bot is unable to exchange token. Ask for user consent."
        );

        await sleep(sleepTimeOffset);

        // Mock Teams sends signin/verifyState message after user consent back to the bot after timeout
        const invokeActivity: Partial<Activity> = createReply(ActivityTypes.Invoke, activity);
        invokeActivity.name = verifyStateOperationName;
        adapter.send(invokeActivity);
      })
      .assertReply((activity) => {
        assert.strictEqual(activity.type, ActivityTypes.Message);
        assert.strictEqual(activity.text, ErrorMessage.FailedToRetrieveSsoToken);
      });
  });

  it("sso execution dialog should only work in MS Teams Channel", async function () {
    const adapter: TestAdapter = await initializeTestEnv(undefined, undefined, Channels.Test);

    await adapter.send("TestCommand").catch((error) => {
      assert.strictEqual(
        error.message,
        "SSO execution dialog is only supported in MS Teams Channel"
      );
    });
  });

  it("sso execution dialog should be able to sign user in and get exchange tokens when consent", async function () {
    this.timeout(500);
    const adapter: TestAdapter = await initializeTestEnv();

    await adapter
      .send("TestCommand")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply((activity) => {
        // User has not consent. Assert bot send out 412
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.PRECONDITION_FAILED);
        assert.strictEqual(
          (activity.value as any).body.failureDetail,
          "The bot is unable to exchange token. Ask for user consent."
        );

        // Mock Teams sends signin/verifyState message after user consent back to the bot
        const invokeActivity: Partial<Activity> = createReply(ActivityTypes.Invoke, activity);
        invokeActivity.name = verifyStateOperationName;
        adapter.send(invokeActivity);
      })
      .assertReply((activity) => {
        // Assert bot send out OAuthCard gain to get SSO token
        assertTeamsSsoOauthCardActivity(activity);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
        assert.strictEqual((activity.value as any).body.id, id);
      })
      .assertReply((activity) => {
        console.log(activity.text);
        assert.strictEqual(activity.type, ActivityTypes.Message);
        assert.strictEqual(activity.text, testSsoHandlerResponseMessage);
      });
  });

  it("sso execution dialog should be able to sign user in and get exchange tokens when consent with regex trigger pattern", async function () {
    this.timeout(500);
    const adapter: TestAdapter = await initializeTestEnv(
      undefined,
      undefined,
      undefined,
      /TestCommand/i
    );

    await adapter
      .send("TestCommand")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply((activity) => {
        // User has not consent. Assert bot send out 412
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.PRECONDITION_FAILED);
        assert.strictEqual(
          (activity.value as any).body.failureDetail,
          "The bot is unable to exchange token. Ask for user consent."
        );

        // Mock Teams sends signin/verifyState message after user consent back to the bot
        const invokeActivity: Partial<Activity> = createReply(ActivityTypes.Invoke, activity);
        invokeActivity.name = verifyStateOperationName;
        adapter.send(invokeActivity);
      })
      .assertReply((activity) => {
        // Assert bot send out OAuthCard gain to get SSO token
        assertTeamsSsoOauthCardActivity(activity);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
        assert.strictEqual((activity.value as any).body.id, id);
      })
      .assertReply((activity) => {
        console.log(activity.text);
        assert.strictEqual(activity.type, ActivityTypes.Message);
        assert.strictEqual(activity.text, testSsoHandlerResponseMessage);
      });
  });

  function createReply(type: ActivityTypes, activity: Partial<Activity>): Partial<Activity> {
    return {
      type: type,
      from: { id: activity.recipient!.id, name: activity.recipient!.name },
      recipient: { id: activity.from!.id, name: activity.from!.name },
      replyToId: activity.id,
      serviceUrl: activity.serviceUrl,
      channelId: activity.channelId,
      conversation: {
        isGroup: activity.conversation!.isGroup,
        id: activity.conversation!.id,
        name: activity.conversation!.name,
        conversationType: "personal",
        tenantId: tenantId,
      },
    };
  }

  function assertTeamsSsoOauthCardActivity(activity: Partial<Activity>): void {
    assert.isArray(activity.attachments);
    assert.strictEqual(activity.attachments?.length, 1);
    assert.strictEqual(activity.attachments?.[0].contentType, CardFactory.contentTypes.oauthCard);
    assert.strictEqual(activity.inputHint, InputHints.AcceptingInput);

    assert.strictEqual(
      (activity.attachments?.[0].content as any).buttons[0].type,
      ActionTypes.Signin
    );
    assert.strictEqual(
      (activity.attachments?.[0].content as any).buttons[0].title,
      "Teams SSO Sign In"
    );

    assert.strictEqual(
      (activity.attachments?.[0].content as any).buttons[0].value,
      `${initiateLoginEndpoint}?scope=${encodeURI(
        requiredScopes.join(" ")
      )}&clientId=${clientId}&tenantId=${tenantId}&loginHint=${userPrincipalName}`
    );
  }

  function mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(
    adapter: TestAdapter,
    activity: Partial<Activity>
  ): void {
    const invokeActivity: Partial<Activity> = createReply(ActivityTypes.Invoke, activity);
    invokeActivity.name = tokenExchangeOperationName;
    invokeActivity.value = {
      id: id,
      token: ssoToken,
    };
    adapter.send(invokeActivity);
  }

  /**
   * Initialize dialogs, adds teamsBotSsoPrompt in dialog set and initialize testAdapter for test case.
   * @param timeout_value positive number set to teamsSsoPromptSettings.timeout property
   * @param endOnInvalidMessage boolean value set to teamsSsoPromptSettings.endOnInvalidMessage property
   * @param channelId value set to dialog context activity channel. Defaults to `Channels.MSteams`.
   */
  async function initializeTestEnv(
    timeout_value?: number,
    endOnInvalidMessage?: boolean,
    channelId?: Channels,
    triggerPatterns: TriggerPatterns = "TestCommand"
  ): Promise<TestAdapter> {
    const storage = new MemoryStorage();

    // Create new ConversationState with MemoryStorage
    const convoState: ConversationState = new ConversationState(new MemoryStorage());

    // Create a DialogState property, DialogSet and TeamsBotSsoPrompt
    const dialogState: AgentStatePropertyAccessor<DialogState> =
      convoState.createProperty("dialogState");
    const dialogs: DialogSet = new DialogSet(dialogState);

    const ssoPromptSettings: TeamsBotSsoPromptSettings = {
      scopes: requiredScopes,
      timeout: timeout_value,
      endOnInvalidMessage: endOnInvalidMessage,
    };
    const ssoExecutionDialog = new BotSsoExecutionDialog(
      storage,
      ssoPromptSettings,
      OnBehalfOfCredentialAuthConfig,
      initiateLoginEndpoint
    );
    const testHandler = new TestSsoCommandHandler(triggerPatterns, testSsoHandlerResponseMessage);
    ssoExecutionDialog.addCommand(
      async (
        context: TurnContext,
        tokenResponse: TeamsBotSsoPromptTokenResponse,
        message: CommandMessage
      ) => {
        const response = await testHandler.handleCommandReceived(context, message, tokenResponse);
        if (typeof response === "string") {
          await context.sendActivity(response);
        } else {
          const replyActivity = response as Partial<Activity>;
          if (replyActivity) {
            await context.sendActivity(replyActivity as Activity);
          }
        }
      },
      testHandler.triggerPatterns
    );
    dialogs.add(ssoExecutionDialog);

    // Initialize TestAdapter.
    const adapter: TestAdapter = new TestAdapter(async (turnContext) => {
      const dc = await dialogs.createContext(turnContext);
      dc.context.activity.channelId = channelId === undefined ? Channels.Msteams : channelId;

      await ssoExecutionDialog.run(turnContext, dialogState);
      await convoState.saveChanges(turnContext, false);
    });

    adapter.onTurnError = async (context: TurnContext, error: Error) => {
      await context.sendActivity(error.message);
    };
    return adapter;
  }
});
