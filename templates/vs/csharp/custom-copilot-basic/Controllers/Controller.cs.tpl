using {{SafeProjectName}}.Models;
using Microsoft.Teams.AI.Models.OpenAI;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Activities.Invokes;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Invokes;
using Microsoft.Teams.Apps.Annotations;

namespace {{SafeProjectName}}.Controllers
{
    [TeamsController]
    public class Controller(OpenAIChatPrompt _prompt)
    {
        [Message]
        public async Task OnMessage(IContext<Microsoft.Teams.Api.Activities.MessageActivity> context, [Context] IContext.Client client)
        {
            var state = State.From(context);

            await _prompt.Send(context.Activity.Text, new() { Messages = state.Messages }, (chunk) => Task.Run(() =>
            {
                context.Stream.Emit(chunk);
            }), context.CancellationToken);

            state.Save(context);

            context.Stream.Emit((Microsoft.Teams.Api.Activities.MessageActivity)new Microsoft.Teams.Api.Activities.MessageActivity().AddFeedback().AddAIGenerated());
        }

        [Microsoft.Teams.Apps.Activities.Invokes.Message.SubmitAction]
        public async Task OnSubmitAction(IContext<Messages.SubmitActionActivity> context)
        {
            Console.WriteLine($"Your feedback is {context.Activity.Value.ActionValue}");

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