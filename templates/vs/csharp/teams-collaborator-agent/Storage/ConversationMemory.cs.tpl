namespace {{SafeProjectName}}.Storage
{
    public class ConversationMemory
    {
        private readonly IDatabase _store;
        private readonly string _conversationId;

        public ConversationMemory(IDatabase store, string conversationId)
        {
            _store = store;
            _conversationId = conversationId;
        }

        public Task AddMessagesAsync(IEnumerable<MessageRecord> messages) =>
            _store.AddMessagesAsync(messages);

        public Task<List<MessageRecord>> ValuesAsync() => _store.GetAsync(_conversationId);

        public Task<int> LengthAsync() => _store.CountMessagesAsync(_conversationId);

        public Task ClearAsync() => _store.ClearConversationAsync(_conversationId);

        public Task<List<MessageRecord>> GetMessagesByTimeRangeAsync(
            string startTime,
            string endTime
        ) => _store.GetMessagesByTimeRangeAsync(_conversationId, startTime, endTime);

        public Task<List<MessageRecord>> GetRecentMessagesAsync(int limit = 10) =>
            _store.GetRecentMessagesAsync(_conversationId, limit);

        public Task<List<MessageRecord>> GetFilteredMessagesAsync(
            string[] keywords,
            string startTime,
            string endTime,
            string[]? participants = null,
            int? maxResults = null
        ) =>
            _store.GetFilteredMessagesAsync(
                _conversationId,
                keywords,
                startTime,
                endTime,
                participants,
                maxResults
            );
    }
}
