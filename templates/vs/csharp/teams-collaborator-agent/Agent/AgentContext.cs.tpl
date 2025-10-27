using {{SafeProjectName}}.Storage;

namespace {{SafeProjectName}}.Agent;

public class AgentContext
{
    public string ConversationId { get; set; } = string.Empty;
    public string UserMessage { get; set; } = string.Empty;
    public DateTime StartTime { get; set; }
    public DateTime? EndTime { get; set; }
    public List<ConversationMember> Members { get; set; } = new();
    public IConversationStorage Storage { get; set; } = null!;
    public string BotId { get; set; } = string.Empty;
}

public class ConversationMember
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Role { get; set; } = "user";
}
