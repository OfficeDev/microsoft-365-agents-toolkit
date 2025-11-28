namespace {{SafeProjectName}}.Storage
{
    public interface IDatabase
    {
        Task InitializeAsync();
        Task ClearAllAsync();
        Task<List<MessageRecord>> GetAsync(string conversationId);
        Task<List<MessageRecord>> GetMessagesByTimeRangeAsync(
            string conversationId,
            string startTime,
            string endTime
        );
        Task<List<MessageRecord>> GetRecentMessagesAsync(string conversationId, int limit = 10);
        Task ClearConversationAsync(string conversationId);
        Task AddMessagesAsync(IEnumerable<MessageRecord> messages);
        Task<int> CountMessagesAsync(string conversationId);
        Task ClearAllMessagesAsync();
        Task<List<MessageRecord>> GetFilteredMessagesAsync(
            string conversationId,
            string[] keywords,
            string startTime,
            string endTime,
            string[]? participants = null,
            int? maxResults = null
        );
        Task<bool> RecordFeedbackAsync(string replyToId, string reaction, object? feedbackJson = null);
        Task CloseAsync();
    }
}
