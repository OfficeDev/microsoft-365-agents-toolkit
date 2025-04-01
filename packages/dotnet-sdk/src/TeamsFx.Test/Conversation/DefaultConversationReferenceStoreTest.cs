namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Core.Models;
    using Microsoft.TeamsFx.Conversation;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Threading.Tasks;

    [Obsolete]
    internal class DefaultConversationReferenceStoreTest
    {
        private readonly InMemoryStorage _storage;
        private readonly DefaultConversationReferenceStore _store;

        public DefaultConversationReferenceStoreTest()
        {
            _storage = new InMemoryStorage();
            _store = new DefaultConversationReferenceStore(_storage);
        }

        [TestMethod]
        public async Task Add_NonExistItem()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            const string key = "key";
            const string activityId = "add_nonexistitem";
            var reference = new ConversationReference
            {
                ActivityId = activityId
            };
            var isAdded = await store.Add(key, reference, new ConversationReferenceStoreAddOptions { });

            Assert.IsTrue(isAdded);
            Assert.Equals(activityId, storage.Items[key].ActivityId);
        }

        [TestMethod]
        public async Task Add_OverwriteExistItem()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            const string key = "key";
            const string oldActivityId = "add_existitem_old";
            const string newActivityId = "add_existitem_new";

            storage.Items[key] = new ConversationReference
            {
                ActivityId = oldActivityId
            };

            var reference = new ConversationReference
            {
                ActivityId = newActivityId
            };
            var isAdded = await store.Add(key, reference, new ConversationReferenceStoreAddOptions { Overwrite = true });

            Assert.IsTrue(isAdded);
            Assert.Equals(newActivityId, storage.Items[key].ActivityId);
        }

        [TestMethod]
        public async Task Add_NotOverwriteExistItem()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            const string key = "key";
            const string oldActivityId = "add_existitem_old";
            const string newActivityId = "add_existitem_new";

            storage.Items[key] = new ConversationReference
            {
                ActivityId = oldActivityId
            };

            var reference = new ConversationReference
            {
                ActivityId = newActivityId
            };
            var isAdded = await store.Add(key, reference, new ConversationReferenceStoreAddOptions { Overwrite = false });

            Assert.IsFalse(isAdded);
            Assert.Equals(oldActivityId, storage.Items[key].ActivityId);
        }

        [TestMethod]
        public async Task Add_NotOverwriteExistItemByDefault()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            const string key = "key";
            const string oldActivityId = "add_existitem_old";
            const string newActivityId = "add_existitem_new";

            storage.Items[key] = new ConversationReference
            {
                ActivityId = oldActivityId
            };

            var reference = new ConversationReference
            {
                ActivityId = newActivityId
            };
            var isAdded = await store.Add(key, reference, new ConversationReferenceStoreAddOptions { });

            Assert.IsFalse(isAdded);
            Assert.Equals(oldActivityId, storage.Items[key].ActivityId);
        }

        [TestMethod]
        public async Task Remove_ExistItem()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            const string key = "key";
            const string activityId = "remove_existitem";

            var reference = new ConversationReference
            {
                ActivityId = activityId
            };
            storage.Items[key] = reference;

            var isRemoved = await store.Remove(key, reference);
            Assert.IsTrue(isRemoved);
            Assert.IsTrue(!storage.Items.ContainsKey(key));
        }

        [TestMethod]
        public async Task Remove_NonExistItem()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            const string key = "key";
            const string activityId = "remove_existitem";

            var reference = new ConversationReference
            {
                ActivityId = activityId
            };

            var isRemoved = await store.Remove(key, reference);
            Assert.IsFalse(isRemoved);
        }

        [TestMethod]
        public async Task List()
        {
            var storage = new InMemoryStorage();
            var store = new DefaultConversationReferenceStore(storage);

            storage.Items.Add("key1", new ConversationReference { ActivityId = "activity-1" });
            storage.Items.Add("key2", new ConversationReference { ActivityId = "activity-2" });
            storage.Items.Add("key3", new ConversationReference { ActivityId = "activity-3" });

            var items = await store.List();
            Assert.AreEqual(3, items.Data.Length);
            Assert.IsNull(items.ContinuationToken);
        }
    }
}
