using Microsoft.Teams.AI;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Apps;

namespace {{SafeProjectName}}.Models
{
    public class State
    {
        public bool Status = false;
        public IList<IMessage> Messages = [];

        public static State From(IContext<IActivity> context)
        {
            return From<IActivity>(context);
        }

        public static State From<TActivity>(IContext<TActivity> context) where TActivity : IActivity
        {
            var key = $"{context.Activity.Conversation.Id}.{context.Activity.From.Id}";
            return (State?)context.Storage.Get(key) ?? new();
        }

        public void Save(IContext<IActivity> context)
        {
            Save<IActivity>(context);
        }

        public void Save<TActivity>(IContext<TActivity> context) where TActivity : IActivity
        {
            var key = $"{context.Activity.Conversation.Id}.{context.Activity.From.Id}";
            context.Storage.Set(key, this);
        }
    }
}