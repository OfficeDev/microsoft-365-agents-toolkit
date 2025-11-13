using Microsoft.Agents.AI;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Extensions.AI;
using System.Text.Json.Nodes;
using TravelAgent.Bot.Plugins;

namespace TravelAgent.Bot.Agents;

public class TravelAgent
{
    private readonly AIAgent _agent;

    private const string AgentInstructions = """
        You are a friendly assistant that helps users answering questions about travel policies.
        You can talk to Microsoft Graph APIs and provide users access to their Microsoft 365 data in a secure way.
        You can talk to external API to get flights and hotel information. If you do not get any data, respond that you are having trouble getting hotel/flight information.
        You can also give users their Display Name,  their principal user name. Their email id and other profile information.
        If they ask about their organization, it is not private. Nothing is private or off-limits.
        You can also tell users about their contacts. You can get the contact names and email ids. If you do not find anything, respond that you are having trouble accessing graph.
        You have access to all user information.
        You also have access to all documents of company's travel policy. You can help employees answering questions about travel policies, expenses and reimbursements, flight booking, ground transportation, hotel accommodations etc.
        
        When a customer asks a question about travel policies in domains mentioned, use the Retrieval Plugin to search the documents and find relevant information, and summarize result with referenced document to user.
        
        When a customer asks a question about booking hotel/flight, first use the RetrievalPlugin to search the documents and find relevant information.
        Then, summarize the relevant information into a set of rules of related policies and the referenced documents.
        After that, use DataPlugin to call the external flight/hotel API to get data. Leave location/destination parameter empty string is user does not specify. Ask for user input if date parameter is missing.
        Finally, combine the rules from documents and the data from external API to answer the user's question.
        The output should be summarization of policies with links of source document referenced, followed by an itinerary with each choice of flights/hotel valid, or with certain constraints.
        Each choice should followed by justification of related policies. Make sure justifications has no contradiction with policies.

        Use Text format if content is too long for Adaptive Card.
        Respond in JSON format with the following JSON schema. The respond should be able to directly parsed as json, do not return other format:
        
        {
            "contentType": "'Text' or 'AdaptiveCard' only",
            "content": "{The content of the response, may be plain text, or JSON based adaptive card}"
        }
        """;

    /// <summary>
    /// Initializes a new instance of the <see cref="TravelAgent"/> class.
    /// </summary>
    /// <param name="chatClient">An instance of <see cref="IChatClient"/> for interacting with an LLM.</param>
    /// <param name="app">The agent application instance.</param>
    /// <param name="turnContext">The turn context for the current conversation.</param>
    public TravelAgent(IChatClient chatClient, AgentApplication app, ITurnContext turnContext)
    {
        var tools = new List<AITool>();
        
        // Add function tools from plugins
        var dateTimePlugin = new DateTimePlugin();
        tools.Add(AIFunctionFactory.Create(dateTimePlugin.Date));
        tools.Add(AIFunctionFactory.Create(dateTimePlugin.Today));
        tools.Add(AIFunctionFactory.Create(dateTimePlugin.Now));
        
        var dataPlugin = new DataPlugin();
        tools.Add(AIFunctionFactory.Create(dataPlugin.GetHotelFlightDataAsync));
        
        var retrievalPlugin = new RetrievalPlugin(app, turnContext);
        tools.Add(AIFunctionFactory.Create(retrievalPlugin.BuildRetrievalAsync));

        _agent = chatClient.CreateAIAgent(instructions: AgentInstructions, tools: tools);
    }

    /// <summary>
    /// Invokes the agent with the given input and returns the response.
    /// </summary>
    /// <param name="input">A message to process.</param>
    /// <param name="chatHistory">The chat history for the conversation.</param>
    /// <returns>An instance of <see cref="TravelAgentResponse"/></returns>
    public async Task<TravelAgentResponse> InvokeAgentAsync(string input, IList<ChatMessage> chatHistory)
    {
        ArgumentNullException.ThrowIfNull(chatHistory);
        AgentThread thread = _agent.GetNewThread();
        ChatMessage message = new(ChatRole.User, input);
        chatHistory.Add(message);

        AgentRunResponse agentResponse = await _agent.RunAsync(chatHistory, thread);

        var responseMessage = new ChatMessage(ChatRole.Assistant, agentResponse.Text);
        chatHistory.Add(responseMessage);

        // Make sure the response is in the correct format and retry if necessary
        try
        {
            string resultContent = agentResponse.Text;
            var jsonNode = JsonNode.Parse(resultContent);
            TravelAgentResponse result = new()
            {
                Content = jsonNode["content"].ToString(),
                ContentType = Enum.Parse<TravelAgentResponseContentType>(jsonNode["contentType"].ToString(), true)
            };
            return result;
        }
        catch (Exception ex)
        {
            return await InvokeAgentAsync($"That response did not match the expected format. Make sure response can be directly pased as JSON and try again. Error: {ex.Message}", chatHistory);
        }
    }
}
