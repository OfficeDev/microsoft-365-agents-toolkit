namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Hosting.AspNetCore;
    using Microsoft.Agents.Hosting.AspNetCore.BackgroundQueue;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;

    [TestClass]
    public class CommandBotTest
    {
        [TestMethod]
        public void CreateCommandBot_ShouldUseMiddleware()
        {
            // Arrange
            var _mockClientFactory = new Mock<IChannelServiceClientFactory>();
            var _mockActivityTaskQueue = new Mock<IActivityTaskQueue>();
            var _mockAdapter = new CloudAdapter(_mockClientFactory.Object, _mockActivityTaskQueue.Object);

            // Act
            CommandBot bot = new CommandBot(_mockAdapter,
                new CommandOptions()
                {
                    Commands = new List<ITeamsCommandHandler> { new TestCommandHandler(("test-command")) }
                });

            // Assert
            Assert.IsNotNull(bot.CommandHandlers);
            Assert.AreEqual(1, bot.CommandHandlers.Count);
            Assert.IsNotNull(_mockAdapter.MiddlewareSet);
        }

        [TestMethod]
        public void RegisterCommand_ShouldSucceed()
        {
            // Arrange
            var _mockClientFactory = new Mock<IChannelServiceClientFactory>();
            var _mockActivityTaskQueue = new Mock<IActivityTaskQueue>();
            var _mockAdapter = new CloudAdapter(_mockClientFactory.Object, _mockActivityTaskQueue.Object);
            CommandBot bot = new CommandBot(_mockAdapter,
                new CommandOptions()
                {
                    Commands = new List<ITeamsCommandHandler> { new TestCommandHandler(("test-command1")) }
                });

            // Act
            bot.RegisterCommand(new TestCommandHandler("test-command2"));

            // Assert
            Assert.IsNotNull(bot.CommandHandlers);
            Assert.AreEqual(2, bot.CommandHandlers.Count);
        }

        [TestMethod]
        public void RegisterCommands_ShouldSucceed()
        {
            // Arrange
            var _mockClientFactory = new Mock<IChannelServiceClientFactory>();
            var _mockActivityTaskQueue = new Mock<IActivityTaskQueue>();
            var _mockAdapter = new CloudAdapter(_mockClientFactory.Object, _mockActivityTaskQueue.Object);
            CommandBot bot = new CommandBot(_mockAdapter, new CommandOptions());
            var Commands = new List<ITeamsCommandHandler>
            { 
                new TestCommandHandler(("test-command1")),
                new TestCommandHandler(("test-command2"))
            };

            // Act
            bot.RegisterCommands(Commands);

            // Assert
            Assert.IsNotNull(bot.CommandHandlers);
            Assert.AreEqual(2, bot.CommandHandlers.Count);
        }
    }
}
