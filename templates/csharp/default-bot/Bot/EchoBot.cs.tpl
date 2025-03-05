using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Extensions.Teams.Compat;

namespace {{SafeProjectName}}.Bot;

public class EchoBot : ActivityHandler
{ 
    protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        string messageText = turnContext.Activity.RemoveRecipientMention()?.Trim();
        var replyText = $"Echo: {messageText}";
        await turnContext.SendActivityAsync(MessageFactory.Text(replyText), cancellationToken);
    }
    protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
    {
        var welcomeText = "Hi there! I'm a Teams bot that will echo what you said to me.";
        foreach (var member in membersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText), cancellationToken);
            }
        }
    }
}

