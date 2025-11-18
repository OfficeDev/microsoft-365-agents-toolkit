using {{SafeProjectName}}.Bot.Agents;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.AI;


namespace {{SafeProjectName}}.Bot;

public class {{SafeProjectName}}Bot : AgentApplication
{
    private Agents.{{SafeProjectName}}Agent _travelAgent;
    private IChatClient _chatClient;

    public {{SafeProjectName}}Bot(AgentApplicationOptions options, IChatClient chatClient) : base(options)
    {
        _chatClient = chatClient ?? throw new ArgumentNullException(nameof(chatClient));

        OnConversationUpdate(ConversationUpdateEvents.MembersAdded, WelcomeMessageAsync);

        // Replace with following line to enable SSO for calling Microsoft 365 Retrieval API
        //OnActivity(ActivityTypes.Message, MessageActivityAsync, rank: RouteRank.Last, autoSignInHandlers: ["graph"]);
        OnActivity(ActivityTypes.Message, MessageActivityAsync, rank: RouteRank.Last);
    }

    protected async Task MessageActivityAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        // Setup local service connection
        ServiceCollection serviceCollection = [
            new ServiceDescriptor(typeof(ITurnState), turnState),
            new ServiceDescriptor(typeof(ITurnContext), turnContext),
            new ServiceDescriptor(typeof(IChatClient), _chatClient),
        ];

        // Start a Streaming Process 
        await turnContext.StreamingResponse.QueueInformativeUpdateAsync("Working on a response for you");

        IList<ChatMessage> chatHistory = turnState.GetValue("conversation.chatHistory", () => new List<ChatMessage>());
        _travelAgent = new Agents.{{SafeProjectName}}Agent(_chatClient, this, turnContext);

        // Invoke the TravelAgent to process the message
        TravelAgentResponse travelResponse = await _travelAgent.InvokeAgentAsync(turnContext.Activity.Text, chatHistory);
        if (travelResponse == null)
        {
            turnContext.StreamingResponse.QueueTextChunk("Sorry, I couldn't get the travel information at the moment.");
            await turnContext.StreamingResponse.EndStreamAsync(cancellationToken);
            return;
        }

        // Create a response message based on the response content type from the TravelAgent
        // Send the response message back to the user. 
        switch (travelResponse.ContentType)
        {
            case TravelAgentResponseContentType.Text:
                turnContext.StreamingResponse.QueueTextChunk(travelResponse.Content);
                break;
            case TravelAgentResponseContentType.AdaptiveCard:
                turnContext.StreamingResponse.FinalMessage = MessageFactory.Attachment(new Attachment()
                {
                    ContentType = "application/vnd.microsoft.card.adaptive",
                    Content = travelResponse.Content,
                });
                break;
            default:
                break;
        }
        await turnContext.StreamingResponse.EndStreamAsync(cancellationToken); // End the streaming response
    }

    protected async Task WelcomeMessageAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        foreach (ChannelAccount member in turnContext.Activity.MembersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Hello and Welcome! I'm here to help with all your travel information needs!"), cancellationToken);
            }
        }
    }
}
