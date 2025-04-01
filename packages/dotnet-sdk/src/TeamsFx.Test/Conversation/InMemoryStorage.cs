namespace Microsoft.TeamsFx.Test.Conversation
{
    using Microsoft.Agents.Core.Models;
    using Microsoft.TeamsFx.Conversation;

    [Obsolete]
    public sealed class InMemoryStorage : INotificationTargetStorage
    {
        public InMemoryStorage()
        {
            Items = new Dictionary<string, ConversationReference>();
        }

        public Dictionary<string, ConversationReference> Items { get; private set; }

        public Task Delete(string key, CancellationToken cancellationToken = default)
        {
            Items.Remove(key);
            return Task.CompletedTask;
        }

        public Task<ConversationReference[]> List(CancellationToken cancellationToken = default)
        {
            return Task.FromResult(Items.Values.ToArray());
        }

        public Task<ConversationReference> Read(string key, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(Items.GetValueOrDefault(key, null));
        }

        public Task Write(string key, ConversationReference reference, CancellationToken cancellationToken = default)
        {
            Items[key] = reference;
            return Task.CompletedTask;
        }
    }
}
