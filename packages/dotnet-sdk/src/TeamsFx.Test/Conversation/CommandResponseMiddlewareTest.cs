namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using System;
    using System.Threading.Tasks;

    [TestClass]
    public class CommandResponseMiddlewareTest
    {
        private readonly CommandResponseMiddleware _middleware;
        private readonly StringTrigger _stringTrigger = new("SampleTest");
        private readonly RegExpTrigger _regexTrigger = new(@"^test (.*?)$");

        public CommandResponseMiddlewareTest()
        {
            _middleware = new CommandResponseMiddleware();
        }

        [DataTestMethod]
        [DataRow("SampleTest")]
        [DataRow("SAMPLETEST")]
        [DataRow(" SampleTest ")]
        public async Task OnTurnAsync_TriggerCommandByString_ShouldSucceed(string input)
        {
            // Arrange
            var command = new TestCommandHandler(new[] { _stringTrigger });
            _middleware.CommandHandlers.Add(command);
            var activity = CreateMessageActivity(input);

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(activity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            Assert.IsTrue(command.IsTriggered);
            Assert.IsNotNull(command.ReceivedMessage);
            Assert.AreEqual(command.ReceivedMessage.Text, input.Trim(), true);
        }

        [DataTestMethod]
        [DataRow("Test fake-value")]
        [DataRow("Test abc123")]
        public async Task OnTurnAsync_TriggerCommandByRegex_ShouldSucceed(string input)
        {
            // Arrange
            var command = new TestCommandHandler(new[] { _regexTrigger });
            _middleware.CommandHandlers.Add(command);
            var param = input.Trim().Split(' ').Last();
            var activity = CreateMessageActivity(input);

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(activity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            Assert.IsTrue(command.IsTriggered);
            Assert.IsNotNull(command.ReceivedMessage);
            Assert.AreEqual(command.ReceivedMessage.Text, input, true);
            Assert.IsNotNull(command.ReceivedMessage.Matches);
            Assert.AreEqual(command.ReceivedMessage.Matches[0].Groups[1].Value, param, true);
        }

        [TestMethod]
        public async Task OnTurnAsync_MultiMatches_TriggerTheFirstMatch()
        {
            // Arrange
            var command1 = new TestCommandHandler(new[] { _regexTrigger });
            var command2 = new TestCommandHandler(new[] { _regexTrigger });
            _middleware.CommandHandlers.Add(command1);
            _middleware.CommandHandlers.Add(command2);

            var input = "Test fake-value";
            var activity = CreateMessageActivity(input);

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(activity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            Assert.IsTrue(command1.IsTriggered);
            Assert.IsFalse(command2.IsTriggered);
        }

        [DataTestMethod]
        [DataRow("test1")]
        [DataRow("test2")]
        public async Task OnTurnAsync_MultiStringPatterns_ShouldTrigger(string input)
        {
            // Arrange
            var command = new TestCommandHandler(new []
            { 
                new StringTrigger("test1"),
                new StringTrigger("test2")
            });

            _middleware.CommandHandlers.Add(command);
            var activity = CreateMessageActivity(input);

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(activity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            Assert.IsTrue(command.IsTriggered);
            Assert.IsNotNull(command.ReceivedMessage);
            Assert.AreEqual(command.ReceivedMessage.Text, input, true);
            Assert.IsTrue(command.IsTriggered);
        }

        [DataTestMethod]
        [DataRow("SampleTest extra-input")]
        [DataRow("Invalid input")]
        [DataRow("A Test !@#$%^")]
        [DataRow(null)]
        public async Task OnTurnAsync_InputNotMatch_ShouldSkipped(string input)
        {
            // Arrange
            var command = new TestCommandHandler(new List<ITriggerPattern> { _stringTrigger, _regexTrigger });
            _middleware.CommandHandlers.Add(command);
            var activity = CreateMessageActivity(input);

            var mockTurnContext = new Mock<ITurnContext>();
            mockTurnContext.Setup(tc => tc.Activity).Returns(activity);

            // Act
            await _middleware.OnTurnAsync(mockTurnContext.Object, (ct) => { return Task.CompletedTask; }, CancellationToken.None);

            // Assert
            Assert.IsFalse(command.IsTriggered);
        }

        private static IActivity CreateMessageActivity(string text)
        {
            return new Activity()
            {
                Type = ActivityTypes.Message,
                Text = text,
                Recipient = new ChannelAccount()
                {
                    Id = Guid.NewGuid().ToString("N")
                }
            };
        }
    }
}
