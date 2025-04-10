// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ActionTypes,
  Activity,
  ActivityTypes,
  CardFactory,
  Channels,
  ConversationState,
  InputHints,
  MemoryStorage,
  AgentStatePropertyAccessor,
  StatusCodes,
} from "@microsoft/agents-hosting";
import {
  TeamsChannelAccount,
  tokenExchangeOperationName,
  verifyStateOperationName,
  TeamsInfo,
} from "@microsoft/agents-hosting-teams";
import { DialogSet, DialogState, DialogTurnStatus } from "@microsoft/agents-hosting-dialogs";
import {
  TeamsBotSsoPrompt,
  TeamsBotSsoPromptTokenResponse,
  OnBehalfOfUserCredential,
  OnBehalfOfCredentialAuthConfig,
  ErrorWithCode,
  ErrorCode,
  TeamsBotSsoPromptSettings,
} from "../../../../src";
import { assert, expect, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import * as sinon from "sinon";
import { AccessToken } from "@azure/identity";
import { promisify } from "util";
import { TestAdapter } from "../conversation/testUtils";

chaiUse(chaiPromises);

describe("TeamsBotSsoPrompt Tests - Node", () => {
  const sleep = promisify(setTimeout);

  const clientId = "fake_client_id";
  const clientSecret = "fake_client_secret";
  const tenantId = "fake_tenant";
  const userPrincipalName = "fake_userPrincipalName";
  const authorityHost = "fake_authority_host";
  const initiateLoginEndpoint = "fake_initiate_login_endpoint";

  const TeamsBotSsoPromptId = "TEAMS_BOT_SSO_PROMPT";
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
  const ssoTokenExpiration = "2018-09-18T01:42:28.000Z";
  const timeoutValue = 50;
  const sleepTimeOffset: number = timeoutValue + 20;
  enum SsoLogInResult {
    Success = "Success",
    Fail = "Fail",
  }
  let getMemberStub: sinon.SinonStub;
  const sandbox = sinon.createSandbox();

  beforeEach(function () {
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

    getMemberStub = sandbox.stub(TeamsInfo, "getMember").callsFake(async () => {
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
  });

  it("teams bot sso prompt should be able to sign user in and get exchange tokens when consent", async function () {
    this.timeout(500);

    const adapter: TestAdapter = await initializeTestEnv();

    await adapter
      .send("Hello")
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
        // Assert bot send out invoke response status 200 to Teams to signal verifivation invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
        assert.strictEqual((activity.value as any).body.id, id);
      })
      .assertReply(SsoLogInResult.Success)
      .assertReply((activity) => {
        // Assert prompt result has exchanged token and sso token.
        const result = JSON.parse(activity.text as string) as TeamsBotSsoPromptTokenResponse;
        assert.strictEqual(result.token, exchangeToken);
        assert.strictEqual(result.ssoToken, ssoToken);
        assert.strictEqual(result.ssoTokenExpiration, ssoTokenExpiration);
      });
  });

  it("teams bot sso prompt should timeout with teams verification invoke activity when wait a long time", async function () {
    const adapter: TestAdapter = await initializeTestEnv(timeoutValue);

    await adapter
      .send("Hello")
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
      .assertReply(SsoLogInResult.Fail);
  });

  it("teams bot sso prompt should timeout with token exchange activity when wait a long time", async function () {
    const adapter: TestAdapter = await initializeTestEnv(timeoutValue);

    await adapter
      .send("Hello")
      .assertReply(async (activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        await sleep(sleepTimeOffset);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply(SsoLogInResult.Fail);
  });

  it("teams bot sso prompt should timeout with message activity when wait a long time", async function () {
    const adapter: TestAdapter = await initializeTestEnv(timeoutValue);

    await adapter
      .send("Hello")
      .assertReply(async (activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        await sleep(sleepTimeOffset);

        // Mock message activity sent to the bot
        const messageActivity: Partial<Activity> = createReply(ActivityTypes.Message, activity);
        messageActivity.text = "message sent to bot.";
        adapter.send(messageActivity);
      })
      .assertReply(SsoLogInResult.Fail);
  });

  it("teams bot sso prompt should end on invalid message when endOnInvalidMessage default to true", async function () {
    const adapter: TestAdapter = await initializeTestEnv(undefined);

    await adapter
      .send("Hello")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock User send invalid message
        const messageActivity = createReply(ActivityTypes.Message, activity);
        messageActivity.text = "user sends invalid message during auth flow";
        adapter.send(messageActivity);
      })
      .assertReply(SsoLogInResult.Fail);
  });

  it("teams bot sso prompt should not end on invalid message when endOnInvalidMessage set to false", async function () {
    const adapter: TestAdapter = await initializeTestEnv(undefined, false);

    await adapter
      .send("Hello")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock User send invalid message, wchich should be ignored.
        const messageActivity = createReply(ActivityTypes.Message, activity);
        messageActivity.text = "user sends invalid message during auth flow";
        adapter.send(messageActivity);

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
        // Assert bot send out invoke response status 200 to Teams to signal verifivation invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual((activity.value as any).status, StatusCodes.OK);
        assert.strictEqual((activity.value as any).body.id, id);
      })
      .assertReply(SsoLogInResult.Success)
      .assertReply((activity) => {
        // Assert prompt result has exchanged token and sso token.
        const result = JSON.parse(activity.text as string) as TeamsBotSsoPromptTokenResponse;
        assert.strictEqual(result.token, exchangeToken);
        assert.strictEqual(result.ssoToken, ssoToken);
        assert.strictEqual(result.ssoTokenExpiration, ssoTokenExpiration);
      });
  });

  it("teams bot sso prompt should only work in MS Teams Channel", async function () {
    const adapter: TestAdapter = await initializeTestEnv(undefined, undefined, Channels.Test);

    await adapter.send("Hello").catch((error) => {
      assert.strictEqual(
        error.message,
        "Teams Bot SSO Prompt is only supported in MS Teams Channel"
      );
    });
  });

  it("teams bot sso prompt should work with undefined user Principal Name", async function () {
    getMemberStub.restore();
    sandbox.stub(TeamsInfo, "getMember").callsFake(async () => {
      const account: TeamsChannelAccount = {
        id: "fake_id",
        name: "fake_name",
        userPrincipalName: "",
      };
      return account;
    });
    const adapter: TestAdapter = await initializeTestEnv();

    await adapter.send("Hello").assertReply((activity) => {
      // Assert bot send out OAuthCard
      assert.strictEqual(
        (activity.attachments?.[0].content as any).buttons[0].value,
        `${initiateLoginEndpoint}?scope=${encodeURI(
          requiredScopes.join(" ")
        )}&clientId=${clientId}&tenantId=${tenantId}&loginHint=`
      );
    });
  });

  it("create TeamsBotSsoPrompt instance should throw InvalidParameter error with invalid scopes", async function () {
    const invalidScopes = [1, 2];
    const settings: any = {
      scopes: invalidScopes,
    };

    expect(() => {
      new TeamsBotSsoPrompt(
        OnBehalfOfCredentialAuthConfig,
        initiateLoginEndpoint,
        TeamsBotSsoPromptId,
        settings
      );
    })
      .to.throw(ErrorWithCode, "The type of scopes is not valid, it must be string or string array")
      .with.property("code", ErrorCode.InvalidParameter);
  });

  it("create TeamsBotSsoPrompt instance should throw InvalidConfiguration error with empty configuration", async function () {
    const settings: any = {
      scopes: requiredScopes,
    };

    expect(() => {
      new TeamsBotSsoPrompt(
        {} as OnBehalfOfCredentialAuthConfig,
        "",
        TeamsBotSsoPromptId,
        settings
      );
    })
      .to.throw(
        ErrorWithCode,
        "clientId, clientSecret or certificateContent, tenantId, authorityHost in configuration is invalid: undefined."
      )
      .with.property("code", ErrorCode.InvalidConfiguration);
  });

  function createReply(type: ActivityTypes, activity: Partial<Activity>): Partial<Activity> {
    return {
      type: type,
      from: { id: activity.recipient?.id as string, name: activity.recipient?.name as string },
      recipient: { id: activity.from?.id as string, name: activity.from?.name as string },
      replyToId: activity.id,
      serviceUrl: activity.serviceUrl,
      channelId: activity.channelId,
      conversation: {
        isGroup: activity.conversation?.isGroup as boolean,
        id: activity.conversation?.id as string,
        name: activity.conversation?.name as string,
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
    channelId?: Channels
  ): Promise<TestAdapter> {
    // Create new ConversationState with MemoryStorage
    const convoState: ConversationState = new ConversationState(new MemoryStorage());

    // Create a DialogState property, DialogSet and TeamsBotSsoPrompt
    const dialogState: AgentStatePropertyAccessor<DialogState> =
      convoState.createProperty("dialogState");
    const dialogs: DialogSet = new DialogSet(dialogState);
    const settings: TeamsBotSsoPromptSettings = {
      scopes: requiredScopes,
      timeout: timeout_value,
      endOnInvalidMessage: endOnInvalidMessage,
    };

    dialogs.add(
      new TeamsBotSsoPrompt(
        OnBehalfOfCredentialAuthConfig,
        initiateLoginEndpoint,
        TeamsBotSsoPromptId,
        settings
      )
    );

    // Initialize TestAdapter.
    const adapter: TestAdapter = new TestAdapter(async (turnContext) => {
      const dc = await dialogs.createContext(turnContext);
      dc.context.activity.channelId = channelId === undefined ? Channels.Msteams : channelId;

      const results = await dc.continueDialog();
      if (results.status === DialogTurnStatus.empty) {
        await dc.beginDialog(TeamsBotSsoPromptId);
      } else if (results.status === DialogTurnStatus.complete) {
        if (results.result?.token) {
          await turnContext.sendActivity(SsoLogInResult.Success);
          const resultStr = JSON.stringify(results.result);
          await turnContext.sendActivity(resultStr);
        } else {
          await turnContext.sendActivity(SsoLogInResult.Fail);
        }
      }
      await convoState.saveChanges(turnContext);
    });
    return adapter;
  }
});
