// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Azure;
using Azure.AI.Agents.Persistent;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using System.Collections.Concurrent;
using Microsoft.Agents.AI;
using Microsoft.Agents.AI.AzureAI;
using System.Text.Json;
using Microsoft.Agents.Core.Serialization;
using Microsoft.Extensions.AI;

namespace {{SafeProjectName}}.Agents;

public class FoundryAgent : AgentApplication
{
    // This is a cache to store the agent model for the Microsoft Foundry agent as this object uses private serializer and virtual objects and is expensive to create.
    // This cache will store the returned model by agent ID. if you need to change the agent model you would need to clear this cache.
    private static ConcurrentDictionary<string, Response<PersistentAgent>> _agentModelCache = new();

    private readonly string _agentId;
    private readonly string _connectionStringForAgent;

    public FoundryAgent(AgentApplicationOptions options, IConfiguration configuration) : base(options)
    {

        // TO DO: get the connection string of your Microsoft Foundry project in the portal
        this._connectionStringForAgent = configuration["AIServices:AzureAIFoundryProjectEndpoint"];
        if (string.IsNullOrEmpty(_connectionStringForAgent))
        {
            throw new InvalidOperationException("AzureAIFoundryProjectEndpoint is not configured.");
        }

        // TO DO: Get the assistant ID in the Microsoft Foundry project portal for your agent
        this._agentId = configuration["AIServices:AgentID"];
        if (string.IsNullOrEmpty(this._agentId))
        {
            throw new InvalidOperationException("AgentID is not configured.");
        }

        // Setup Agent with Route handlers to manage connecting and responding from the Microsoft Foundry agent

        // This is handling the sign out event, which will clear the user authorization token.
        OnMessage("--signout", HandleSignOutAsync);

        // This is handling the clearing of the agent model cache without needing to restart the agent.
        OnMessage("--clearcache", HandleClearingModelCacheAsync);

        // This is handling the message activity, which will send the user message to the Microsoft Foundry agent.
        // we are also indicating which auth profile we want to have available for this handler.
        OnActivity(ActivityTypes.Message, SendMessageToFoundryAgent, autoSignInHandlers: ["SSO"]);
    }

    /// <summary>
    /// Handle the clearing of the agent model cache.
    /// </summary>
    /// <param name="turnContext"></param>
    /// <param name="turnState"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    private async Task HandleClearingModelCacheAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        _agentModelCache.Clear();
        await turnContext.SendActivityAsync("The agent model cache has been cleared.", cancellationToken: cancellationToken);
        Console.WriteLine("The agent model cache has been cleared.");
    }

    /// <summary>
    /// Handle the sign out event, and clear the logged in user token
    /// </summary>
    /// <param name="turnContext"></param>
    /// <param name="turnState"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    private async Task HandleSignOutAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        await UserAuthorization.SignOutUserAsync(turnContext, turnState, cancellationToken: cancellationToken);
        await turnContext.SendActivityAsync("You have signed out", cancellationToken: cancellationToken);
    }

    /// <summary>
    /// This method sends the user message ( just text in this example ) to the Microsoft Foundry agent and streams the response back to the user.
    /// </summary>
    /// <param name="turnContext"></param>
    /// <param name="turnState"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    protected async Task SendMessageToFoundryAgent(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        Console.WriteLine($"\nUser message received: {turnContext.Activity.Text}\n");
        try
        {
            // Start a Streaming Process to let clients that support streaming know that we are processing the request.
            await turnContext.StreamingResponse.QueueInformativeUpdateAsync("Just a moment please..", cancellationToken).ConfigureAwait(false);

            // Set up the PersistentAgentsClient to communicate with the Microsoft Foundry agent.

            // This is a helper class to generate an OBO User Token for the Microsoft Foundry agent from the current user authorization.
            PersistentAgentsClient _aiProjectClient = new PersistentAgentsClient(this._connectionStringForAgent,
                        // This is a helper class to generate an OBO User Token for the Microsoft Foundry agent from the current user authorization.
                        new UserAuthorizationTokenWrapper(UserAuthorization, turnContext, "SSO"));

            // Get the requested agent by ID.
            Response<PersistentAgent> agentModel = _agentModelCache.TryGetValue(this._agentId, out var cachedModel) ? cachedModel : null;
            if (agentModel == null)
            {
                // subtle hint to the client that the agent model is being fetched.
                await turnContext.StreamingResponse.QueueInformativeUpdateAsync("Connecting to Microsoft Foundry.", cancellationToken).ConfigureAwait(false);

                // If the agent model is not found in the conversation state, fetch it from the Microsoft Foundry project.
                agentModel = await _aiProjectClient.Administration.GetAgentAsync(this._agentId).ConfigureAwait(false);
                // Cache the agent model for future use.
                _agentModelCache.TryAdd(this._agentId, agentModel);
            }

            // Create an instance of the AIAgent using the new GetAIAgentAsync extension method
            // The new API accepts agentId (string) instead of the Response<PersistentAgent> object
            AIAgent _existingAgent = await _aiProjectClient.GetAIAgentAsync(this._agentId, cancellationToken: cancellationToken).ConfigureAwait(false);

            // Get or create thread:
            AgentThread _agentThread = GetConversationThread(_existingAgent, turnState);

            // Inform the client that we are working on a response
            await turnContext.StreamingResponse.QueueInformativeUpdateAsync("Sending request to Foundry Agent.", cancellationToken).ConfigureAwait(false);

            // Create a new message to send to the Azure agent
            ChatMessage message = new(ChatRole.User, turnContext.Activity.Text);
            // Send the message to the Azure agent and get the response
            // This will handle text responses,  if you want to handle attachments and other content types, you would need to extend this method.
            await foreach (AgentRunResponseUpdate response in _existingAgent.RunStreamingAsync(message, _agentThread, cancellationToken: cancellationToken))
            {
                if (!string.IsNullOrEmpty(response.Text))
                {
                    turnContext.StreamingResponse.QueueTextChunk(response.Text);
                }
            }
            turnState.Conversation.SetValue("conversation.threadInfo", ProtocolJsonSerializer.ToJson(_agentThread.Serialize()));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending message to Azure agent: {ex.Message}");
            turnContext.StreamingResponse.QueueTextChunk($"An error occurred while processing your request. {ex.Message}");
        }
        finally
        {
            await turnContext.StreamingResponse.EndStreamAsync(cancellationToken).ConfigureAwait(false); // End the streaming response
        }
    }


    /// <summary>
    /// Manage Agent threads against the conversation state.
    /// </summary>
    /// <param name="agent">ChatAgent</param>
    /// <param name="turnState">State Manager for the Agent.</param>
    /// <returns></returns>
    private static AgentThread GetConversationThread(AIAgent agent, ITurnState turnState)
    {
        AgentThread thread;
        string agentThreadInfo = turnState.Conversation.GetValue<string>("conversation.threadInfo", () => null);
        if (string.IsNullOrEmpty(agentThreadInfo))
        {
            thread = agent.GetNewThread();
        }
        else
        {
            JsonElement ele = ProtocolJsonSerializer.ToObject<JsonElement>(agentThreadInfo);
            thread = agent.DeserializeThread(ele);
        }
        return thread;
    }
}
