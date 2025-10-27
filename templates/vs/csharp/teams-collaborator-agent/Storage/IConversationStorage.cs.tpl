using {{SafeProjectName}}.Storage.Models;

namespace {{SafeProjectName}}.Storage;

public interface IConversationStorage
{
    Task InitializeAsync();
    Task<bool> AddMessageAsync(ConversationMessage message);
    Task<bool> AddMessagesAsync(IEnumerable<ConversationMessage> messages);
    Task<IEnumerable<ConversationMessage>> GetMessagesByTimeRangeAsync(
        string conversationId, DateTime startTime, DateTime? endTime = null);
    Task<IEnumerable<ConversationMessage>> GetAllMessagesAsync(string conversationId);
    Task<bool> ClearConversationAsync(string conversationId);
    Task<bool> RecordFeedbackAsync(string messageId, string reaction, string? feedbackJson = null);
}
