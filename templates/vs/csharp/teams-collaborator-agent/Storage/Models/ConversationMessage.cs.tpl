namespace {{SafeProjectName}}.Storage.Models;

public class ConversationMessage
{
    public string Id { get; set; } = string.Empty;
    public string ConversationId { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; }
    public string Role { get; set; } = "user";
}

public class MessageFeedback
{
    public string MessageId { get; set; } = string.Empty;
    public string Reaction { get; set; } = string.Empty;
    public string? FeedbackJson { get; set; }
    public DateTime Timestamp { get; set; }
}
