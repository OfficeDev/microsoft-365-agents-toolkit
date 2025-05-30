using Microsoft.Teams.AI.State;

namespace {{SafeProjectName}}.Model
{
    // Extend the turn state by configuring custom strongly typed state classes.
    public class AppState : TurnState
    {
        public AppState()
        {
            ScopeDefaults[CONVERSATION_SCOPE] = new ConversationState();
        }

        /// <summary>
        /// Stores all the conversation-related state.
        /// </summary>
        public new ConversationState Conversation
        {
            get
            {
                TurnStateEntry? scope = GetScope(CONVERSATION_SCOPE);

                if (scope == null)
                {
                    throw new ArgumentException("TurnState hasn't been loaded. Call LoadStateAsync() first.");
                }

                return (ConversationState)scope.Value!;
            }
            set
            {
                TurnStateEntry? scope = GetScope(CONVERSATION_SCOPE);

                if (scope == null)
                {
                    throw new ArgumentException("TurnState hasn't been loaded. Call LoadStateAsync() first.");
                }

                scope.Replace(value!);
            }
        }
    }

    public class MyTask
    {
        public string Title { get; set; }
        public string Description { get; set; }
    }


    // This class adds custom properties to the turn state which will be accessible in the various handler methods.
    public class ConversationState : Record
    {
        public Dictionary<string, MyTask> Tasks
        {
            get => Get<Dictionary<string, MyTask>>("tasks") ?? new Dictionary<string, MyTask>();
            set => Set("tasks", value);
        }
    }
}
