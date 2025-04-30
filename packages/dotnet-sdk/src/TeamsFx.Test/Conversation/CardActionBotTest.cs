namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Hosting.AspNetCore;
    using Microsoft.Agents.Hosting.AspNetCore.BackgroundQueue;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using System.Collections.Generic;

    [TestClass]
    public class CardActionBotTest
    {
        [TestMethod]
        public void CreateCardActionBot_ShouldUseMiddleware()
        {
            // Arrange
            var _mockClientFactory = new Mock<IChannelServiceClientFactory>();
            var _mockActivityTaskQueue = new Mock<IActivityTaskQueue>();
            var _mockAdapter = new CloudAdapter(_mockClientFactory.Object, _mockActivityTaskQueue.Object);

            // Act
            var bot = new CardActionBot(_mockAdapter,
                new CardActionOptions()
                {
                    Actions = new List<IAdaptiveCardActionHandler> { new Mock<IAdaptiveCardActionHandler>().Object }
                });

            // Assert
            Assert.IsNotNull(bot.CardActionHandlers);
            Assert.AreEqual(1, bot.CardActionHandlers.Count);
            Assert.IsNotNull(_mockAdapter.MiddlewareSet);
        }

        [TestMethod]
        public void RegisterHandler_ShouldSucceed()
        {
            // Arrange
            var _mockClientFactory = new Mock<IChannelServiceClientFactory>();
            var _mockActivityTaskQueue = new Mock<IActivityTaskQueue>();
            var _mockAdapter = new CloudAdapter(_mockClientFactory.Object, _mockActivityTaskQueue.Object);
            var bot = new CardActionBot(_mockAdapter, new CardActionOptions());

            // Act
            bot.RegisterHandler(new Mock<IAdaptiveCardActionHandler>().Object);

            // Assert
            Assert.IsNotNull(bot.CardActionHandlers);
            Assert.AreEqual(1, bot.CardActionHandlers.Count);
        }

        [TestMethod]
        public void RegisterHandlers_ShouldSucceed()
        {
            // Arrange
            var _mockClientFactory = new Mock<IChannelServiceClientFactory>();
            var _mockActivityTaskQueue = new Mock<IActivityTaskQueue>();
            var _mockAdapter = new CloudAdapter(_mockClientFactory.Object, _mockActivityTaskQueue.Object);
            var bot = new CardActionBot(_mockAdapter, new CardActionOptions());
            var handlers = new List<IAdaptiveCardActionHandler>
            {
                new Mock<IAdaptiveCardActionHandler>().Object,
                new Mock<IAdaptiveCardActionHandler>().Object
            };

            // Act
            bot.RegisterHandlers(handlers);

            // Assert
            Assert.IsNotNull(bot.CardActionHandlers);
            Assert.AreEqual(2, bot.CardActionHandlers.Count);
        }
    }
}
