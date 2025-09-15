﻿using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Utils;
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
            var text = TextUtils.StripMentionsText(context.Activity);
            
            var additionalContext = dataSource.RenderData(text);
            var enrichedText = $"{text}\n\nAdditional Context: {additionalContext}";

            if (context.Activity.Conversation.IsGroup == true)
            {
                var response = await _prompt.Send(enrichedText, new() { Messages = state.Messages }, null, context.CancellationToken);
                await context.Send(new Microsoft.Teams.Api.Activities.MessageActivity(response.Content).AddFeedback().AddAIGenerated());
            }
            else
            {
                await _prompt.Send(enrichedText, new() { Messages = state.Messages }, (chunk) => Task.Run(() =>
                {
                    context.Stream.Emit(chunk);
                }), context.CancellationToken);

                context.Stream.Emit((Microsoft.Teams.Api.Activities.MessageActivity)new Microsoft.Teams.Api.Activities.MessageActivity().AddFeedback().AddAIGenerated());
            }
            state.Save(context);
        }

        [Microsoft.Teams.Apps.Activities.Invokes.Message.SubmitAction]
        public Task OnSubmitAction(IContext<Messages.SubmitActionActivity> context)
        {
            Console.WriteLine($"Your feedback is {context.Activity.Value.ActionValue}");
            return Task.CompletedTask;
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