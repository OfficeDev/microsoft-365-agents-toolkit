namespace Microsoft.TeamsFx.Test.Conversation
{
    using System.Text.Json;
    using Microsoft.Agents.Core.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class ConversationReferenceExtensionsTest
    {
        [TestMethod]
        public void Clone_IsDeepClone()
        {
            var source = new ConversationReference
            {
                ActivityId = "x",
                ChannelId = "y",
                Conversation = new ConversationAccount
                {
                    Id = "foo",
                    Name = "bar",
                },
            };

            var target = source.Clone();
            Assert.AreEqual(JsonSerializer.Serialize(source), JsonSerializer.Serialize(target));

            source.ActivityId = "xx";
            source.Conversation.Id = "foofoo";
            Assert.AreEqual("x", target.ActivityId);
            Assert.AreEqual("foo", target.Conversation.Id);
        }

        [TestMethod]
        public void GetKey_HappyPath()
        {
            var ref1 = new ConversationReference
            {
                ActivityId = "x",
                ChannelId = "y",
                Conversation = new ConversationAccount
                {
                    Id = "foo",
                    Name = "bar",
                    TenantId = "a",
                },
            };
            var ref2 = new ConversationReference
            {
                ActivityId = "x",
                ChannelId = "y",
            };
            var ref3 = new ConversationReference
            {
                ActivityId = "x",
                ChannelId = "y",
                Conversation = new ConversationAccount
                {
                    Id = "foo",
                    Name = "bar",
                },
            };

            var key1 = ref1.GetKey();
            var key2 = ref2.GetKey();
            var key3 = ref3.GetKey();

            Assert.AreEqual("_a_foo", key1);
            Assert.AreEqual("__", key2);
            Assert.AreEqual("__foo", key3);
        }

        [DataTestMethod]
        [DataRow("personal", NotificationTargetType.Person)]
        [DataRow("PERSONAL", NotificationTargetType.Person)]
        [DataRow("groupChat", NotificationTargetType.Group)]
        [DataRow("GROUPcHAT", NotificationTargetType.Group)]
        [DataRow("channel", NotificationTargetType.Channel)]
        [DataRow("CHANNEL", NotificationTargetType.Channel)]
        [DataRow("", NotificationTargetType.Unknown)]
        [DataRow("x", NotificationTargetType.Unknown)]
        [DataRow(null, NotificationTargetType.Unknown)]
        public void GetTargetType_HappyPath(string conversationType, NotificationTargetType expectedType)
        {
            var reference = new ConversationReference
            {
                Conversation = new ConversationAccount
                {
                    ConversationType = conversationType,
                },
            };

            var actualType = reference.GetTargetType();

            Assert.AreEqual(expectedType, actualType);
        }
    }
}