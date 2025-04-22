namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using Moq;

    [TestClass]
    public class ITurnContextExtensionsTest
    {
        [TestMethod]
        public void GetTeamsBotInstallationId_NoChannelData()
        {
            var activity = new Activity();
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(ctx => ctx.Activity).Returns(activity);

            var id = mockContext.Object.GetTeamsBotInstallationId();

            Assert.IsNull(id);
        }

        [TestMethod]
        public void GetTeamsBotInstallationId_IsTeam()
        {
            var activity = new Activity
            {
                ChannelData = new TeamsChannelData
                {
                    Team = new TeamInfo
                    {
                        Id = "foo",
                    },
                },
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(ctx => ctx.Activity).Returns(activity);

            var id = mockContext.Object.GetTeamsBotInstallationId();

            Assert.AreEqual("foo", id);
        }

        [TestMethod]
        public void GetTeamsBotInstallationId_NotTeam()
        {
            var activity = new Activity
            {
                Conversation = new ConversationAccount
                {
                    Id = "foo",
                }
            };
            var mockContext = new Mock<ITurnContext>();
            mockContext.SetupGet(ctx => ctx.Activity).Returns(activity);

            var id = mockContext.Object.GetTeamsBotInstallationId();

            Assert.AreEqual("foo", id);
        }
    }
}
