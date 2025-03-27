namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.BotBuilder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using Moq;

    [TestClass]
    [Obsolete]
    public class NotificationMiddlewareTest
    {
        private readonly InMemoryStorage _storage;
        private readonly NotificationMiddleware _middleware;

        public NotificationMiddlewareTest()
        {
            _storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(_storage);
            _middleware = new NotificationMiddleware(store);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotInstalled()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "installationUpdate",
                Action = "add",
                Id = activityId,
                ChannelId = "x",
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(1, _storage.Items.Count);
            var reference = _storage.Items.GetValueOrDefault($"_a_{conversationId}", null);
            Assert.IsNotNull(reference);
            Assert.AreEqual(activityId, reference.ActivityId);
            Assert.AreEqual("x", reference.ChannelId);
            Assert.AreEqual("a", reference.Conversation.TenantId);
            Assert.AreEqual(conversationId, reference.Conversation.Id);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotUninstalled()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "installationUpdate",
                Action = "remove",
                Id = activityId,
                ChannelId = "x",
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                },
            };
            var reference = activity.GetConversationReference();
            await _storage.Write(reference.GetKey(), reference, CancellationToken.None);

            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(1, _storage.Items.Count);
            var existingReference = _storage.Items.GetValueOrDefault($"_a_{conversationId}", null);
            Assert.IsNotNull(reference);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(0, _storage.Items.Count);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotMessaged_Personal()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "message",
                Id = activityId,
                ChannelId = "x",
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                    ConversationType = "personal",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(1, _storage.Items.Count);
            var reference = _storage.Items.GetValueOrDefault($"_a_{conversationId}", null);
            Assert.IsNotNull(reference);
            Assert.AreEqual(activityId, reference.ActivityId);
            Assert.AreEqual("x", reference.ChannelId);
            Assert.AreEqual("a", reference.Conversation.TenantId);
            Assert.AreEqual(conversationId, reference.Conversation.Id);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotMessaged_GroupChat()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "message",
                Id = activityId,
                ChannelId = "x",
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                    ConversationType = "groupChat",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(1, _storage.Items.Count);
            var reference = _storage.Items.GetValueOrDefault($"_a_{conversationId}", null);
            Assert.IsNotNull(reference);
            Assert.AreEqual(activityId, reference.ActivityId);
            Assert.AreEqual("x", reference.ChannelId);
            Assert.AreEqual("a", reference.Conversation.TenantId);
            Assert.AreEqual(conversationId, reference.Conversation.Id);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotMessaged_Channel()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var teamId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "message",
                Id = activityId,
                ChannelId = "x",
                ChannelData = new TeamsChannelData
                {
                    Team = new TeamInfo
                    {
                        Id = teamId,
                    },
                },
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                    ConversationType = "channel",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(1, _storage.Items.Count);
            var reference = _storage.Items.GetValueOrDefault($"_a_{teamId}", null);
            Assert.IsNotNull(reference);
            Assert.AreEqual(activityId, reference.ActivityId);
            Assert.AreEqual("x", reference.ChannelId);
            Assert.AreEqual("a", reference.Conversation.TenantId);
            Assert.AreEqual(teamId, reference.Conversation.Id);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotMessaged_GeneralChannel()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var teamId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "message",
                Id = activityId,
                ChannelId = "x",
                ChannelData = new TeamsChannelData
                {
                    Team = new TeamInfo
                    {
                        Id = teamId,
                    },
                    Channel = new ChannelInfo
                    {
                        Id = teamId,
                    },
                },
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                    ConversationType = "channel",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(1, _storage.Items.Count);
            var reference = _storage.Items.GetValueOrDefault($"_a_{teamId}", null);
            Assert.IsNotNull(reference);
            Assert.AreEqual(activityId, reference.ActivityId);
            Assert.AreEqual("x", reference.ChannelId);
            Assert.AreEqual("a", reference.Conversation.TenantId);
            Assert.AreEqual(teamId, reference.Conversation.Id);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleCurrentBotMessaged_NonGeneralChannel()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var teamId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "message",
                Id = activityId,
                ChannelId = "x",
                ChannelData = new TeamsChannelData
                {
                    Team = new TeamInfo
                    {
                        Id = teamId,
                    },
                    Channel = new ChannelInfo
                    {
                        Id = $"{teamId}-channel",
                    },
                },
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                    ConversationType = "channel",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(0, _storage.Items.Count);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleTeamRestored()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "conversationUpdate",
                Id = activityId,
                ChannelId = "x",
                ChannelData = new TeamsChannelData
                {
                    EventType = "teamRestored",
                },
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(0, _storage.Items.Count);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(1, _storage.Items.Count);
            var reference = _storage.Items.GetValueOrDefault($"_a_{conversationId}", null);
            Assert.IsNotNull(reference);
            Assert.AreEqual(activityId, reference.ActivityId);
            Assert.AreEqual("x", reference.ChannelId);
            Assert.AreEqual("a", reference.Conversation.TenantId);
            Assert.AreEqual(conversationId, reference.Conversation.Id);
        }

        [TestMethod]
        public async Task OnTurnAsync_HandleTeamDeleted()
        {
            var activityId = Guid.NewGuid().ToString("N");
            var conversationId = Guid.NewGuid().ToString("N");
            var activity = new Activity
            {
                Type = "conversationUpdate",
                Id = activityId,
                ChannelId = "x",
                ChannelData = new TeamsChannelData
                {
                    EventType = "teamDeleted",
                },
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                    Name = "bar",
                    TenantId = "a",
                },
            };
            var reference = activity.GetConversationReference();
            await _storage.Write(reference.GetKey(), reference, CancellationToken.None);

            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(c => c.Activity).Returns(activity);

            Assert.AreEqual(1, _storage.Items.Count);
            var existingReference = _storage.Items.GetValueOrDefault($"_a_{conversationId}", null);
            Assert.IsNotNull(reference);
            await _middleware.OnTurnAsync(mockContext.Object, (ctx) => Task.CompletedTask, CancellationToken.None);

            Assert.AreEqual(0, _storage.Items.Count);
        }
    }
}
