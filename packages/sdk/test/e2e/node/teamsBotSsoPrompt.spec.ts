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
  TestAdapter,
} from "@microsoft/agents-hosting";
import {
  TeamsChannelAccount,
  tokenExchangeOperationName,
  TeamsInfo,
} from "@microsoft/agents-hosting-teams";
import { DialogSet, DialogState, DialogTurnStatus } from "@microsoft/agents-hosting-dialogs";
import {
  TeamsBotSsoPrompt,
  TeamsBotSsoPromptTokenResponse,
  TeamsBotSsoPromptSettings,
  OnBehalfOfCredentialAuthConfig,
} from "../../../src";
import { assert, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import { extractIntegrationEnvVariables, getSsoTokenFromTeams } from "../helper";
import { parseJwt } from "../../../src/util/utils";
import * as sinon from "sinon";

chaiUse(chaiPromises);
extractIntegrationEnvVariables();

describe("TeamsBotSsoPrompt Tests with Auth Config - Node", () => {
  const initiateLoginEndpoint = "fake_initiate_login_endpoint";
  const userPrincipalName = "fake_userPrincipalName";

  const TeamsBotSsoPromptId = "TEAMS_BOT_SSO_PROMPT";
  const requiredScopes: string[] = ["User.Read"];
  const invokeResponseActivityType = "invokeResponse";
  const id = "fake_id";

  let ssoToken: string;
  enum SsoLogInResult {
    Success = "Success",
    Fail = "Fail",
  }
  const sandbox = sinon.createSandbox();

  before(async function () {
    ssoToken = await getSsoTokenFromTeams();
    sandbox.stub(TeamsInfo, "getMember").callsFake(async () => {
      const account: TeamsChannelAccount = {
        id: "fake_id",
        name: "fake_name",
        userPrincipalName: userPrincipalName,
      };
      return account;
    });
  });

  after(function () {
    sandbox.restore();
  });

  it("teams bot sso prompt should not be able to sign user in and get exchange tokens when not consent", async function () {
    this.timeout(5000);

    const notConsentScopes = ["Calendars.Read"];
    const adapter: TestAdapter = await initializeTestEnv({ scopes: notConsentScopes });

    await adapter
      .send("Hello")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity, notConsentScopes);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply((activity) => {
        // User has not consent. Assert bot send out 412
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual(activity.value.status, StatusCodes.PRECONDITION_FAILED);
        assert.strictEqual(
          activity.value.body.failureDetail,
          "The bot is unable to exchange token. Ask for user consent."
        );
      });
  });

  it("teams bot sso prompt should be able to sign user in and get exchange tokens when consent", async function () {
    this.timeout(5000);

    const adapter: TestAdapter = await initializeTestEnv({});

    await adapter
      .send("Hello")
      .assertReply((activity) => {
        // Assert bot send out OAuthCard
        assertTeamsSsoOauthCardActivity(activity);

        // Mock Teams sends signin/tokenExchange message with SSO token back to the bot
        mockTeamsSendsTokenExchangeInvokeActivityWithSsoToken(adapter, activity);
      })
      .assertReply((activity) => {
        // Assert bot send out invoke response status 200 to Teams to signal verifivation invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual(activity.value.status, StatusCodes.OK);
      })
      .assertReply(SsoLogInResult.Success)
      .assertReply((activity) => {
        // Assert prompt result has exchanged token and sso token.
        const result = JSON.parse(activity.text as string) as TeamsBotSsoPromptTokenResponse;
        assert.strictEqual(result.ssoToken, ssoToken);
        const accessTokenObj = parseJwt(result.token);
        const ssoTokenObj = parseJwt(result.ssoToken);
        assert.strictEqual(accessTokenObj.oid, ssoTokenObj.oid);
      });
  });

  it("teams bot sso prompt should not end on invalid message when endOnInvalidMessage set to false", async function () {
    const adapter: TestAdapter = await initializeTestEnv({ endOnInvalidMessage: false });

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
        // Assert bot send out invoke response status 200 to Teams to signal token response request invoke has been received
        assert.strictEqual(activity.type, invokeResponseActivityType);
        assert.strictEqual(activity.value.status, StatusCodes.OK);
        assert.strictEqual(activity.value.body.id, id);
      })
      .assertReply(SsoLogInResult.Success)
      .assertReply((activity) => {
        // Assert prompt result has exchanged token and sso token.
        const result = JSON.parse(activity.text as string) as TeamsBotSsoPromptTokenResponse;
        assert.strictEqual(result.ssoToken, ssoToken);
        const accessTokenObj = parseJwt(result.token);
        const ssoTokenObj = parseJwt(result.ssoToken);
        assert.strictEqual(accessTokenObj.oid, ssoTokenObj.oid);
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
        tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID,
      },
    };
  }

  function assertTeamsSsoOauthCardActivity(
    activity: Partial<Activity>,
    scopes: string[] = ["User.Read"]
  ): void {
    assert.isArray(activity.attachments);
    assert.strictEqual(activity.attachments?.length, 1);
    assert.strictEqual(activity.attachments![0].contentType, CardFactory.contentTypes.oauthCard);
    assert.strictEqual(activity.inputHint, InputHints.AcceptingInput);

    assert.strictEqual(activity.attachments![0].content.buttons[0].type, ActionTypes.Signin);
    assert.strictEqual(activity.attachments![0].content.buttons[0].title, "Teams SSO Sign In");
    assert.strictEqual(
      activity.attachments![0].content.buttons[0].value,
      `${initiateLoginEndpoint}?scope=${encodeURI(scopes.join(" "))}&clientId=${
        process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID
      }&tenantId=${process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID}&loginHint=${userPrincipalName}`
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
  async function initializeTestEnv(param: InitializeParams): Promise<TestAdapter> {
    // Create new ConversationState with MemoryStorage
    const convoState: ConversationState = new ConversationState(new MemoryStorage());

    // Create a DialogState property, DialogSet and TeamsBotSsoPrompt
    const dialogState: AgentStatePropertyAccessor<DialogState> =
      convoState.createProperty("dialogState");
    const dialogs: DialogSet = new DialogSet(dialogState);

    let botScopes = param.scopes;
    if (!botScopes) {
      botScopes = requiredScopes;
    }

    const settings: TeamsBotSsoPromptSettings = {
      scopes: botScopes,
      timeout: param.timeout_value,
      endOnInvalidMessage: param.endOnInvalidMessage,
    };

    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const loginUrl = "fake_initiate_login_endpoint";
    dialogs.add(new TeamsBotSsoPrompt(authConfig, loginUrl, TeamsBotSsoPromptId, settings));

    // Initialize TestAdapter.
    const adapter: TestAdapter = new TestAdapter(async (turnContext) => {
      const dc = await dialogs.createContext(turnContext);
      dc.context.activity.channelId =
        param.channelId === undefined ? Channels.Msteams : param.channelId;

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

interface InitializeParams {
  scopes?: string[];
  timeout_value?: number;
  endOnInvalidMessage?: boolean;
  channelId?: Channels;
}
