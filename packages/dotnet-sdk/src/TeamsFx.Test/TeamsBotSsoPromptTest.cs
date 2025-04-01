// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.BotBuilder.Compat;
using Microsoft.Agents.BotBuilder.Dialogs;
using Microsoft.Agents.BotBuilder.State;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Extensions.Teams.Models;
using Microsoft.Agents.Storage;
using Microsoft.Identity.Client;
using Microsoft.TeamsFx.Bot;
using Microsoft.TeamsFx.Configuration;
using Microsoft.TeamsFx.Helper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json.Linq;
using System.Text.Json;

namespace Microsoft.TeamsFx.Test;

public static class SsoResult
{
    public const string Success = "Success";
    public const string Fail = "Fail";
};

[TestClass]
public class TeamsBotSsoPromptTest
{
    private static TeamsBotSsoPromptSettings teamsBotSsoPromptSettingsMock;
    private static readonly string testClientId = Guid.NewGuid().ToString();
    private static readonly string testClientSecret = Guid.NewGuid().ToString();
    private static readonly string testTenantId = Guid.NewGuid().ToString();
    private static readonly string testApplicationIdUri = "fake_application_id_url";
    private static readonly string testOAuthAuthority = $"https://login.microsoftonline.com/{testTenantId}";
    private static readonly string testInitiateLoginEndpoint  = "https://fake_bot_domain/bot-auth-start";
    private static readonly string testDialogId = "TEST_TEAMS_BOT_SSO_PROMPT";
    private const string testName = "test_name";
    private const string testUserId = "test_user_id";
    private const string testUserPrincipalName = "test_user_principal_name";
    /// <summary>
    /// Fake sso token payload, debug using https://jwt.io/
    /// {
    ///   "oid": "fake-oid",
    ///   "name": "fake-name",
    ///   "ver": "1.0",
    ///   "exp": 1893456001,
    ///   "upn": "fake-upn",
    ///   "tid": "fake-tid",
    ///   "aud": "fake-aud"
    /// }
    /// </summary>
    private readonly string fakeAccessToken = "eyJhbGciOiJIUzI1NiJ9.eyJvaWQiOiJmYWtlLW9pZCIsIm5hbWUiOiJmYWtlLW5hbWUiLCJ2ZXIiOiIxLjAiLCJleHAiOjE4OTM0NTYwMDEsInVwbiI6ImZha2UtdXBuIiwidGlkIjoiZmFrZS10aWQiLCJhdWQiOiJmYWtlLWF1ZCJ9.koLfiJSWCFbDXgWV7cauoXtuswTW80MIxLsp2oomaCk";

    /// <summary>
    /// Fake sso token payload, debug using https://jwt.io/
    /// {
    ///   "oid": "fake-oid",
    ///   "name": "fake-name",
    ///   "ver": "1.0",
    ///   "exp": 1893456000, // 2030/1/1
    ///   "upn": "fake-upn",
    ///   "tid": "fake-tid",
    ///   "aud": "fake-aud"
    /// }
    /// </summary>
    private readonly string fakeSsoToken = "eyJhbGciOiJIUzI1NiJ9.eyJvaWQiOiJmYWtlLW9pZCIsIm5hbWUiOiJmYWtlLW5hbWUiLCJ2ZXIiOiIxLjAiLCJleHAiOjE4OTM0NTYwMDAsInVwbiI6ImZha2UtdXBuIiwidGlkIjoiZmFrZS10aWQiLCJhdWQiOiJmYWtlLWF1ZCJ9.IpOpEOoAoqVShYafBEGPr9w8dxYPRU9aln5YRBvoajE";
    private readonly DateTime expiration = new DateTime(2030, 1, 1);
    private static readonly BotAuthenticationOptions botAuthOptions = new BotAuthenticationOptions
    {
        ClientId = testClientId,
        ClientSecret = testClientSecret,
        ApplicationIdUri = testApplicationIdUri,
        OAuthAuthority = testOAuthAuthority,
        InitiateLoginEndpoint  = testInitiateLoginEndpoint 
    };
    private static readonly string[] scopes = new string[] { "User.Read" };

    [ClassInitialize]
    public static void TestFixtureSetup(TestContext _)
    {
        // Executes once for the test class. (Optional)
        teamsBotSsoPromptSettingsMock = new TeamsBotSsoPromptSettings(botAuthOptions, scopes);
    }

    #region ConstructTeamsBotSsoPrompt
    [TestMethod]
    public void TeamsBotSsoPromptWithEmptyDialogIdShouldFail()
    {
        var ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(string.Empty, teamsBotSsoPromptSettingsMock));
        Assert.AreEqual(ExceptionCode.InvalidParameter, ex.Code);
        Assert.AreEqual("Parameter dialogId is null or empty.", ex.Message);
    }

    [TestMethod]
    public void TeamsBotSsoPromptWithEmptySettingShouldFail()
    {
        var ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(testDialogId, null));
        Assert.AreEqual(ExceptionCode.InvalidParameter, ex.Code);
        Assert.AreEqual("Parameter settings is null or empty.", ex.Message);
    }
    #endregion

    #region BeginDialog
    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogWithNoDialogContextShouldFail()
    {
        var ex = await Assert.ThrowsExceptionAsync<ExceptionWithCode>(async () =>
        {
            var prompt = new TeamsBotSsoPrompt(testDialogId, teamsBotSsoPromptSettingsMock);
            await prompt.BeginDialogAsync(null);
        });
        Assert.AreEqual(ExceptionCode.InvalidParameter, ex.Code);
        Assert.AreEqual("Parameter dialogContext is null or empty.", ex.Message);
    }

    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogNotInTeamsShouldFail()
    {
        var ex = await Assert.ThrowsExceptionAsync<ExceptionWithCode>(async () =>
        {
            var prompt = new TeamsBotSsoPrompt(testDialogId, teamsBotSsoPromptSettingsMock);
            var convoState = new ConversationState(new MemoryStorage());
            var dialogState = convoState.CreateProperty<DialogState>("dialogState");

            var adapter = new TestAdapter()
                .Use(new AutoSaveStateMiddleware(convoState));

            // Create new DialogSet.
            var dialogs = new DialogSet(dialogState);
            dialogs.Add(prompt);

            using (var tc = new TurnContext(adapter, new Activity() { Type = ActivityTypes.Message, Conversation = new ConversationAccount() { Id = testUserId }, ChannelId = "not-teams" }))
            {
                var dc = await dialogs.CreateContextAsync(tc);

                await prompt.BeginDialogAsync(dc);
            }
        });
        Assert.AreEqual(ExceptionCode.ChannelNotSupported, ex.Code);
    }

    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogShouldSuccess()
    {
        var ccaMock = new Mock<IIdentityClientAdapter>();
        var mockAuthenticationResult = new AuthenticationResult(fakeAccessToken, true, null, expiration, expiration, string.Empty, null, null, null, Guid.Empty);
        ccaMock.Setup(_ => _.GetAccessToken(It.IsAny<string>(), It.IsAny<IEnumerable<string>>())).ReturnsAsync(mockAuthenticationResult);
        var teamsInfoMock = new Mock<ITeamsInfo>();
        teamsInfoMock.Setup(_ => _.GetTeamsMemberAsync(It.IsAny<ITurnContext>(), It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync(new TeamsChannelAccount
        {
            Id = testUserId,
            Name = testName,
            UserPrincipalName = testUserPrincipalName,
            TenantId = testTenantId
        });
        
        var convoState = new ConversationState(new MemoryStorage());
        var dialogState = convoState.CreateProperty<DialogState>("dialogState");

        // Create new DialogSet.
        var dialogs = new DialogSet(dialogState);
        var prompt = new TeamsBotSsoPrompt(testDialogId, teamsBotSsoPromptSettingsMock);
        prompt._identityClientAdapter = ccaMock.Object;
        prompt._teamsInfo = teamsInfoMock.Object;
        dialogs.Add(prompt);

        var testFlow = InitTestFlow(convoState, dialogs);

        await testFlow
        .Send(new Activity()
        {
            ChannelId = Channels.Msteams,
            Text = "hello",
            Conversation = new ConversationAccount() { Id = testUserId }
        })
        .AssertReply(activity =>
        {
            Assert.AreEqual(1, ((Activity)activity).Attachments.Count);
            Assert.AreEqual(OAuthCard.ContentType, ((Activity)activity).Attachments[0].ContentType);
            OAuthCard card = ((Activity)activity).Attachments[0].Content as OAuthCard;
            Assert.IsNotNull(card);
            Assert.AreEqual(1, card!.Buttons.Count);
            Assert.AreEqual(ActionTypes.Signin, card!.Buttons[0].Type);
            Assert.AreEqual($"{testInitiateLoginEndpoint }?scope=User.Read&clientId={testClientId}&tenantId={testTenantId}&loginHint={testUserPrincipalName}", card!.Buttons[0].Value);
            Assert.AreEqual($"{testApplicationIdUri}/access_as_user", card!.TokenExchangeResource.Uri);
        })
        .Send(new Activity()
        {
            ChannelId = Channels.Msteams,
            Type = ActivityTypes.Invoke,
            Name = SignInConstants.TokenExchangeOperationName,
            Value = JObject.FromObject(new TokenExchangeInvokeRequest()
            {
                Id = "fake_id",
                Token = fakeSsoToken
            })
        })
        .AssertReply(a =>
        {
            Assert.AreEqual(Microsoft.Bot.Schema.ActivityTypesEx.InvokeResponse, a.Type);
            var response = ((Activity)a).Value as InvokeResponse;
            Assert.IsNotNull(response);
            Assert.AreEqual(200, response!.Status);
        })
        .AssertReply(SsoResult.Success)
        .AssertReply(activity =>
        {
            var response = JsonSerializer.Deserialize<TeamsBotSsoPromptTokenResponse>(((Activity)activity).Text);
            Assert.AreEqual(fakeSsoToken, response.SsoToken);
            var expectedSsoExpiration = DateTimeOffset.FromUnixTimeSeconds(long.Parse("1893456000"));
            Assert.AreEqual(expectedSsoExpiration.ToString(), response.SsoTokenExpiration);
            Assert.AreEqual(fakeAccessToken, response.Token);
            Assert.AreEqual(mockAuthenticationResult.ExpiresOn.ToString(), response.Expiration);
        })
        .StartTestAsync();
    }

    [TestMethod]
    public async Task TeamsBotSsoPromptWithEmptyScopeShouldReturnSsoToken()
    {
        var teamsInfoMock = new Mock<ITeamsInfo>();
        teamsInfoMock.Setup(_ => _.GetTeamsMemberAsync(It.IsAny<ITurnContext>(), It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync(new TeamsChannelAccount
        {
            Id = testUserId,
            Name = testName,
            UserPrincipalName = testUserPrincipalName,
            TenantId = testTenantId
        });

        var convoState = new ConversationState(new MemoryStorage());
        var dialogState = convoState.CreateProperty<DialogState>("dialogState");

        // Create new DialogSet.
        var dialogs = new DialogSet(dialogState);

        // Create TeamsBotSsoPrompt with empty scope
        var prompt = new TeamsBotSsoPrompt(testDialogId, new TeamsBotSsoPromptSettings(botAuthOptions, new string[] { }));
        prompt._teamsInfo = teamsInfoMock.Object;
        dialogs.Add(prompt);

        var testFlow = InitTestFlow(convoState, dialogs);

        await testFlow
        .Send(new Activity()
        {
            ChannelId = Channels.Msteams,
            Text = "hello",
            Conversation = new ConversationAccount() { Id = testUserId }
        })
        .AssertReply(activity =>
        {
            Assert.AreEqual(1, ((Activity)activity).Attachments.Count);
            Assert.AreEqual(OAuthCard.ContentType, ((Activity)activity).Attachments[0].ContentType);
            OAuthCard card = ((Activity)activity).Attachments[0].Content as OAuthCard;
            Assert.IsNotNull(card);
            Assert.AreEqual(1, card!.Buttons.Count);
            Assert.AreEqual(ActionTypes.Signin, card!.Buttons[0].Type);
            Assert.AreEqual($"{testInitiateLoginEndpoint }?scope=&clientId={testClientId}&tenantId={testTenantId}&loginHint={testUserPrincipalName}", card!.Buttons[0].Value);
            Assert.AreEqual($"{testApplicationIdUri}/access_as_user", card!.TokenExchangeResource.Uri);
        })
        .Send(new Activity()
        {
            ChannelId = Channels.Msteams,
            Type = ActivityTypes.Invoke,
            Name = SignInConstants.TokenExchangeOperationName,
            Value = JObject.FromObject(new TokenExchangeInvokeRequest()
            {
                Id = "fake_id",
                Token = fakeSsoToken
            })
        })
        .AssertReply(a =>
        {
            Assert.AreEqual(Microsoft.Bot.Schema.ActivityTypesEx.InvokeResponse, a.Type);
            var response = ((Activity)a).Value as InvokeResponse;
            Assert.IsNotNull(response);
            Assert.AreEqual(200, response!.Status);
        })
        .AssertReply(SsoResult.Success)
        .AssertReply(activity =>
        {
            var response = JsonSerializer.Deserialize<TeamsBotSsoPromptTokenResponse>(((Activity)activity).Text);
            Assert.AreEqual(fakeSsoToken, response.SsoToken);
            var expectedSsoExpiration = DateTimeOffset.FromUnixTimeSeconds(long.Parse("1893456000"));
            Assert.AreEqual(expectedSsoExpiration.ToString(), response.SsoTokenExpiration);
            Assert.AreEqual(fakeSsoToken, response.Token);
            Assert.AreEqual(expectedSsoExpiration.ToString(), response.Expiration);
        })
        .StartTestAsync();
    }

    private TestFlow InitTestFlow(ConversationState convoState, DialogSet dialogs)
    {
        var adapter = new TestAdapter()
            .Use(new AutoSaveStateMiddleware(convoState));

        BotCallbackHandler botCallbackHandler = async (turnContext, cancellationToken) =>
        {
            var dc = await dialogs.CreateContextAsync(turnContext, cancellationToken);

            var results = await dc.ContinueDialogAsync(cancellationToken);
            if (results.Status == DialogTurnStatus.Empty)
            {
                await dc.PromptAsync(testDialogId, new PromptOptions(), cancellationToken: cancellationToken);
            }
            else if (results.Status == DialogTurnStatus.Complete)
            {
                if (results.Result is TokenResponse)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(SsoResult.Success), cancellationToken);
                    await turnContext.SendActivityAsync(MessageFactory.Text(JsonSerializer.Serialize(results.Result)), cancellationToken);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(SsoResult.Fail), cancellationToken);
                }
            }
        };

        return new TestFlow(adapter, botCallbackHandler);
    }
    #endregion
}
