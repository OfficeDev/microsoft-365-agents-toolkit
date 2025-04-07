using Microsoft.Agents.BotBuilder.State;

namespace {{SafeProjectName}};

public static class ConversationStateExtensions
{
    public static int MessageCount(this ConversationState state) => state.GetValue<int>("countKey");

    public static void MessageCount(this ConversationState state, int value) => state.SetValue("countKey", value);

    public static int IncrementMessageCount(this ConversationState state)
    {
        int count = state.GetValue<int>("countKey");
        state.SetValue("countKey", ++count);
        return count;
    }
}
