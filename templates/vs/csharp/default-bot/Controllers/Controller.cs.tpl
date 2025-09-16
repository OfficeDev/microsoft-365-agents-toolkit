using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Annotations;

namespace {{SafeProjectName}}.Controllers
{
    [TeamsController]
    public class Controller()
    {
        [Message]
        public async Task OnMessage([Context] MessageActivity activity, [Context] IContext.Client client, [Context] Microsoft.Teams.Common.Logging.ILogger log)
        {
            log.Info("hit!");
            await client.Typing();
            await client.Send($"you said '{activity.Text}'");
        }

        [Conversation.MembersAdded]
        public async Task OnMembersAdded(IContext<ConversationUpdateActivity> context)
        {
            var welcomeText = "How can I help you today?";
            foreach (var member in context.Activity.MembersAdded)
            {
                if (member.Id != context.Activity.Recipient.Id)
                {
                    await context.Send(welcomeText);
                }
            }

        }
    }
}