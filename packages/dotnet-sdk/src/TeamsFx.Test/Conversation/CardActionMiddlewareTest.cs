namespace Microsoft.TeamsFx.Test.Conversation
{
    using AdaptiveCards;
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    [TestClass]
    public class CardActionMiddlewareTest
    {
        private CardActionMiddleware _middleware;

        [TestInitialize]
        public void Initialize()
        {
            _middleware = new CardActionMiddleware();
        }

        [TestMethod]
        public async Task OnTurnAsync_VerbMatch_ShouldTriggerHandler()
        {
            // Arrange
            var mockHandler = CreateMockCardActionHandler("doStuff", InvokeResponseFactory.TextMessage("sample response"));

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            mockHandler.Verify(
                m => m.HandleActionInvokedAsync(
                    It.IsAny<ITurnContext>(), It.IsAny<object>(), It.IsAny<CancellationToken>()),
                Times.Once);
        }

        [TestMethod]
        public async Task OnTurnAsync_VerbMultipleMatch_OnlyTriggerTheFirstHandler()
        {
            // Arrange
            var mockHandler1 = CreateMockCardActionHandler("doStuff", InvokeResponseFactory.TextMessage("sample response"));
            var mockHandler2 = CreateMockCardActionHandler("doStuff", InvokeResponseFactory.TextMessage("sample response"));

            _middleware.CardActionHandlers.Add(mockHandler1.Object);
            _middleware.CardActionHandlers.Add(mockHandler2.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            mockHandler1.Verify(
                m => m.HandleActionInvokedAsync(
                    It.IsAny<ITurnContext>(), It.IsAny<object>(), It.IsAny<CancellationToken>()),
                Times.Once);
            mockHandler2.Verify(
                m => m.HandleActionInvokedAsync(
                    It.IsAny<ITurnContext>(), It.IsAny<object>(), It.IsAny<CancellationToken>()),
                Times.Never);
        }

        [TestMethod]
        public async Task OnTurnAsync_VerbNotMatch_ShouldNotTriggerHandler()
        {
            // Arrange
            var mockHandler = CreateMockCardActionHandler("doStuff", InvokeResponseFactory.TextMessage("sample response"));

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("invalid-verb");

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            mockHandler.Verify(
                m => m.HandleActionInvokedAsync(
                    It.IsAny<ITurnContext>(), It.IsAny<object>(), It.IsAny<CancellationToken>()),
                Times.Never);
        }

        [TestMethod]
        public async Task OnTurnAsync_AdaptiveCard_ReplaceForInteractor_ShouldSendResponseSuccessfully()
        {
            // Arrange
            var mockHandler = CreateMockCardActionHandler(
                "doStuff",
                InvokeResponseFactory.AdaptiveCard(GetAdaptiveCard()),
                AdaptiveCardResponse.ReplaceForInteractor);

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var activites = new List<Activity>();
            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);
            mockTurnContext
                .Setup(tc => tc.SendActivityAsync(It.IsAny<IActivity>(), It.IsAny<CancellationToken>()))
                .Callback<IActivity, CancellationToken>((activity, token) => { activites.Add(activity as Activity); });

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            var expectedActivityCount = 1; // invoke response sent in middleware
            Assert.AreEqual(expectedActivityCount, activites.Count);

            var invokeResponse = activites[0].Value as InvokeResponse;
            Assert.IsNotNull(invokeResponse);
            Assert.IsTrue(invokeResponse.IsSuccessStatusCode());

            var responseBody = invokeResponse.Body as AdaptiveCardInvokeResponse;
            Assert.IsNotNull(responseBody);
            Assert.AreEqual(InvokeResponseContentType.AdaptiveCard, responseBody.Type);
            Assert.IsNotNull(responseBody.Value);
            Assert.AreEqual(GetAdaptiveCard().ToString(), responseBody.Value.ToString());
        }

        [TestMethod]
        public async Task OnTurnAsync_AdaptiveCard_ReplaceForAll_ShouldSendResponseSuccessfully()
        {
            // Arrange
            var mockHandler = CreateMockCardActionHandler(
                "doStuff",
                InvokeResponseFactory.AdaptiveCard(GetAdaptiveCard()),
                AdaptiveCardResponse.ReplaceForAll);

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var activites = new List<Activity>();
            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);
            mockTurnContext
                .Setup(tc => tc.SendActivityAsync(It.IsAny<IActivity>(), It.IsAny<CancellationToken>()))
                .Callback<IActivity, CancellationToken>((activity, token) => { activites.Add(activity as Activity); });
            mockTurnContext
                .Setup(tc => tc.UpdateActivityAsync(It.IsAny<IActivity>(), It.IsAny<CancellationToken>()))
                .Callback<IActivity, CancellationToken>((activity, token) => { activites.Add(activity as Activity); });

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            var expectedActivityCount = 2;   // invoke response and updated activity sent in middleware
            Assert.AreEqual(expectedActivityCount, activites.Count);

            var invokeResponse = activites[0].Value as InvokeResponse;
            Assert.IsNotNull(invokeResponse);
            Assert.IsTrue(invokeResponse.IsSuccessStatusCode());

            var responseBody = invokeResponse.Body as AdaptiveCardInvokeResponse;
            Assert.IsNotNull(responseBody);
            Assert.AreEqual(InvokeResponseContentType.AdaptiveCard, responseBody.Type);
            Assert.IsNotNull(responseBody.Value);
            Assert.AreEqual(GetAdaptiveCard().ToString(), responseBody.Value.ToString());
        }

        [TestMethod]
        public async Task OnTurnAsync_AdaptiveCard_NewForAll_ShouldSendResponseSuccessfully()
        {
            // Arrange
            var mockHandler = CreateMockCardActionHandler(
                "doStuff",
                InvokeResponseFactory.AdaptiveCard(GetAdaptiveCard()),
                AdaptiveCardResponse.NewForAll);

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var activites = new List<Activity>();
            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);
            mockTurnContext
                .Setup(tc => tc.SendActivityAsync(It.IsAny<IActivity>(), It.IsAny<CancellationToken>()))
                .Callback<IActivity, CancellationToken>((activity, token) => { activites.Add(activity as Activity); });

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            var expectedActivityCount = 2;   // default invoke response and new message activity sent in middleware
            Assert.AreEqual(expectedActivityCount, activites.Count);

            var invokeResponse = activites[0].Value as InvokeResponse;
            Assert.IsNotNull(invokeResponse);
            Assert.IsTrue(invokeResponse.IsSuccessStatusCode());

            var responseBody = invokeResponse.Body as AdaptiveCardInvokeResponse;
            Assert.IsNotNull(responseBody);
            Assert.AreEqual(InvokeResponseContentType.Message, responseBody.Type);
            Assert.IsNotNull(responseBody.Value);
            Assert.AreEqual(_middleware.defaultMessage, responseBody.Value);
        }

        [TestMethod]
        public async Task OnTurnAsync_TextMessage_ShouldSendResponseSuccessfully()
        {
            // Arrange
            var testMessage = "This is a sample response";
            var mockHandler = CreateMockCardActionHandler("doStuff", InvokeResponseFactory.TextMessage(testMessage));

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var activites = new List<Activity>();
            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);
            mockTurnContext
                .Setup(tc => tc.SendActivityAsync(It.IsAny<IActivity>(), It.IsAny<CancellationToken>()))
                .Callback<IActivity, CancellationToken>((activity, token) => { activites.Add(activity as Activity); });

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            var expectedActivityCount = 1;   // invoke response sent in middleware
            Assert.AreEqual(expectedActivityCount, activites.Count);

            var invokeResponse = activites[0].Value as InvokeResponse;
            Assert.IsNotNull(invokeResponse);
            Assert.IsTrue(invokeResponse.IsSuccessStatusCode());

            var responseBody = invokeResponse.Body as AdaptiveCardInvokeResponse;
            Assert.IsNotNull(responseBody);
            Assert.AreEqual(InvokeResponseContentType.Message, responseBody.Type);
            Assert.IsNotNull(responseBody.Value);
            Assert.AreEqual(testMessage, responseBody.Value);
        }

        [TestMethod]
        public async Task OnTurnAsync_ErrorResponse_ShouldSendResponseSuccessfully()
        {
            // Arrange
            var testMessage = "This is an error response";
            var mockHandler = CreateMockCardActionHandler("doStuff", InvokeResponseFactory.ErrorResponse(InvokeResponseErrorCode.BadRequest, testMessage));

            _middleware.CardActionHandlers.Add(mockHandler.Object);
            var invokeActivity = CreateInvokeActivity("doStuff");

            var activites = new List<Activity>();
            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(invokeActivity);
            mockTurnContext
                .Setup(tc => tc.SendActivityAsync(It.IsAny<IActivity>(), It.IsAny<CancellationToken>()))
                .Callback<IActivity, CancellationToken>((activity, token) => { activites.Add(activity as Activity); });

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            var expectedActivityCount = 1;   // invoke response sent in middleware
            Assert.AreEqual(expectedActivityCount, activites.Count);

            var invokeResponse = activites[0].Value as InvokeResponse;
            Assert.IsNotNull(invokeResponse);
            Assert.IsTrue(invokeResponse.IsSuccessStatusCode());

            var responseBody = invokeResponse.Body as AdaptiveCardInvokeResponse;
            Assert.IsNotNull(responseBody);
            Assert.AreEqual(InvokeResponseContentType.Error, responseBody.Type);
            Assert.AreEqual((int)InvokeResponseErrorCode.BadRequest, responseBody.StatusCode);

            var botError = responseBody.Value as Error;
            Assert.IsNotNull(botError);
            Assert.AreEqual(testMessage, botError.Message);
            Assert.AreEqual(InvokeResponseErrorCode.BadRequest.ToString(), botError.Code);
        }

        private static Activity CreateInvokeActivity(string verb, object data = null)
        {
            var value = JObject.FromObject(
                new AdaptiveCardInvokeValue
                {
                    Action = new AdaptiveCardInvokeAction
                    {
                        Type = "Action.Execute",
                        Verb = verb,
                        Data = data
                    }
                });

            var activity = new Activity
            {
                Type = ActivityTypes.Invoke,
                Name = "adaptiveCard/action",
                Value = value
            };

            return activity;
        }

        private static Mock<IAdaptiveCardActionHandler> CreateMockCardActionHandler(string verb, InvokeResponse response, AdaptiveCardResponse adaptiveCardResponse = AdaptiveCardResponse.ReplaceForInteractor)
        {
            var mockHandler = new Mock<IAdaptiveCardActionHandler>();
            mockHandler.Setup(_ => _.TriggerVerb).Returns(verb);
            mockHandler.Setup(_ => _.AdaptiveCardResponse).Returns(adaptiveCardResponse);
            mockHandler.Setup(_ => _.HandleActionInvokedAsync(
                It.IsAny<ITurnContext>(),
                It.IsAny<object>(),
                It.IsAny<CancellationToken>())
            ).ReturnsAsync(response);

            return mockHandler;
        }

        private static object GetAdaptiveCard()
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0));

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = "This is a sample response card.",
            });

            return card;
        }
    }
}
