using Microsoft.SemanticKernel.Connectors.OpenAI;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Agents;
using System.Threading.Tasks;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Text;
using System.Text.Json;
using System;
using {{SafeProjectName}}.Bot.Plugins;

namespace {{SafeProjectName}}.Bot.Agents;

public class WeatherForecastAgent
{
    private readonly Kernel _kernel;
    private readonly ChatCompletionAgent _agent;
    private int retryCount;

    private const string AgentName = "WeatherForecastAgent";
    private const string AgentInstructions = """
        You are a friendly assistant that helps people find a weather forecast for a given time and place.
        You may ask follow up questions until you have enough informatioon to answer the customers question,
        but once you have a forecast forecast, make sure to format it nicely using an adaptive card.

        Respond in JSON format with the following JSON schema:
        
        {
            "contentType": "'Text' or 'AdaptiveCard' only",
            "content": "{The content of the response, may be plain text, or JSON based adaptive card}"
        }
        """;

    /// <summary>
    /// Initializes a new instance of the <see cref="WeatherForecastAgent"/> class.
    /// </summary>
    /// <param name="kernel">An instance of <see cref="Kernel"/> for interacting with an LLM.</param>
    public WeatherForecastAgent(Kernel kernel)
    {
        _kernel = kernel;

        // Define the agent
        _agent =
            new()
            {
                Instructions = AgentInstructions,
                Name = AgentName,
                Kernel = _kernel,
                Arguments = new KernelArguments(new OpenAIPromptExecutionSettings() 
                { 
                    FunctionChoiceBehavior = FunctionChoiceBehavior.Auto(), 
                    ResponseFormat = "json_object" 
                }),
            };

        // Give the agent some tools to work with
        _agent.Kernel.Plugins.Add(KernelPluginFactory.CreateFromType<DateTimePlugin>());
        _agent.Kernel.Plugins.Add(KernelPluginFactory.CreateFromType<WeatherForecastPlugin>());
        _agent.Kernel.Plugins.Add(KernelPluginFactory.CreateFromType<AdaptiveCardPlugin>());
    }

    /// <summary>
    /// Invokes the agent with the given input and returns the response.
    /// </summary>
    /// <param name="input">A message to process.</param>
    /// <returns>An instance of <see cref="WeatherForecastAgentResponse"/></returns>
    public async Task<WeatherForecastAgentResponse> InvokeAgentAsync(string input, ChatHistory chatHistory)
    {
        ArgumentNullException.ThrowIfNull(chatHistory);

        ChatMessageContent message = new(AuthorRole.User, input);
        chatHistory.Add(message);

        StringBuilder sb = new();
        await foreach (ChatMessageContent response in _agent.InvokeAsync(chatHistory))
        {
            chatHistory.Add(response);
            sb.Append(response.Content);
        }

        // Make sure the response is in the correct format and retry if neccesary
        try
        {
            var resultContent = sb.ToString();
            var result = JsonSerializer.Deserialize<WeatherForecastAgentResponse>(resultContent);
            retryCount = 0;
            return result;
        }
        catch (JsonException je)
        {
            // Limit the number of retries
            if (retryCount > 2)
            {
                throw;
            }

            // Try again, providing corrective feedback to the model so that it can correct its mistake
            retryCount++;
            return await InvokeAgentAsync($"That response did not match the expected format. Please try again. Error: {je.Message}", chatHistory);
        }
    }
}
