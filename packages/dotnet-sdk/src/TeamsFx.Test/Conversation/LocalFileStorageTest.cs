namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Core.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [Obsolete]
    public class LocalFileStorageTest
    {
        private const string testDir = "./test-local";

        [TestInitialize]
        public void TestInitialize()
        {
            Directory.CreateDirectory(testDir);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Directory.Delete(testDir, true);
        }

        [TestMethod]
        public async Task List_EmptyFile()
        {
            File.WriteAllText(Path.Combine(testDir, ".notification.localstore.json"), "");
            var storage = new LocalFileStorage(testDir);
            var list = await storage.List(CancellationToken.None);

            Assert.IsNotNull(list);
            Assert.AreEqual(0, list.Length);
        }

        [TestMethod]
        public async Task List_NoFile()
        {
            File.Delete(Path.Combine(testDir, ".notification.localstore.json"));
            var storage = new LocalFileStorage(testDir);
            var list = await storage.List(CancellationToken.None);

            Assert.IsNotNull(list);
            Assert.AreEqual(0, list.Length);
        }

        [TestMethod]
        public async Task Add_NonExistingItem()
        {
            var storage = new LocalFileStorage(testDir);
            await storage.Write("key-1", new ConversationReference { ActivityId = "activity-1" });

            var list = await storage.List(CancellationToken.None);
            Assert.IsNotNull(list);
            Assert.AreEqual(1, list.Length);
            Assert.AreEqual("activity-1", list[0].ActivityId);
        }

        [TestMethod]
        public async Task Add_ExistingItem()
        {
            var storage = new LocalFileStorage(testDir);
            await storage.Write("key-1", new ConversationReference { ActivityId = "activity-1" });
            await storage.Write("key-1", new ConversationReference { ActivityId = "activity-2" });

            var list = await storage.List(CancellationToken.None);
            Assert.IsNotNull(list);
            Assert.AreEqual(1, list.Length);
            Assert.AreEqual("activity-2", list[0].ActivityId);
        }

        [TestMethod]
        public async Task Set_FileNameWithEnvironmentVariable()
        {
            var previous = Environment.GetEnvironmentVariable("TEAMSFX_NOTIFICATION_STORE_FILENAME");
            Environment.SetEnvironmentVariable("TEAMSFX_NOTIFICATION_STORE_FILENAME", ".notification.testtool.json");
            var storage = new LocalFileStorage(testDir);
            await storage.Write("key-1", new ConversationReference { ActivityId = "activity-1" });
            Environment.SetEnvironmentVariable("TEAMSFX_NOTIFICATION_STORE_FILENAME", previous);

            var storeFile = Path.Combine(testDir, ".notification.testtool.json");
            Assert.IsTrue(File.Exists(storeFile));

        }
    }
}