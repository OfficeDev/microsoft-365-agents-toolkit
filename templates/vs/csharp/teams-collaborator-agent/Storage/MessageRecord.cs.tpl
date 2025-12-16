namespace {{SafeProjectName}}.Storage
{
    public class MessageRecord
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string ConversationId { get; set; } = string.Empty;
        public string UserId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public string ActivityId { get; set; } = string.Empty;
        public string Timestamp { get; set; } = DateTime.UtcNow.ToString("o");
        public Dictionary<string, object>? Metadata { get; set; }
    }
}
