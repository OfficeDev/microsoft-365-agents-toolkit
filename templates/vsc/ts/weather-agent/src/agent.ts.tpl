import { HumanMessage, SystemMessage } from "@langchain/core/messages";
import { MemorySaver } from "@langchain/langgraph";
import { createReactAgent } from "@langchain/langgraph/prebuilt";
import { AzureChatOpenAI, ChatOpenAI } from "@langchain/openai";
import { ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplicationBuilder,
  MessageFactory,
  TurnContext,
} from "@microsoft/agents-hosting";
import { dateTool } from "./tools/dateTimeTool";
import { getWeatherTool } from "./tools/getWeatherTool";

export const weatherAgent = new AgentApplicationBuilder().build();

weatherAgent.conversationUpdate(
  "membersAdded",
  async (context: TurnContext) => {
    await context.sendActivity(
      `Hello and Welcome! I'm here to help with all your weather forecast needs!`
    );
  }
);

interface WeatherForecastAgentResponse {
  contentType: "Text" | "AdaptiveCard";
  content: string;
}

{{#useOpenAI}}
const agentModel = new ChatOpenAI({
  apiKey: process.env.OPENAI_API_KEY,
  model: "gpt-3.5-turbo",
  temperature: 0,
});
{{/useOpenAI}}
{{#useAzureOpenAI}}
const agentModel = new AzureChatOpenAI({
  azureOpenAIApiVersion: "2024-12-01-preview",
  azureOpenAIApiKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  azureOpenAIApiDeploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
  temperature: 0,
});
{{/useAzureOpenAI}}

const agentTools = [getWeatherTool, dateTool];
const agentCheckpointer = new MemorySaver();
const agent = createReactAgent({
  llm: agentModel,
  tools: agentTools,
  checkpointSaver: agentCheckpointer,
});

const sysMessage = new SystemMessage(`
You are a friendly assistant that helps people find a weather forecast for a given time and place.
You may ask follow up questions until you have enough informatioon to answer the customers question,
but once you have a forecast forecast, make sure to format it nicely using an adaptive card.

Respond in JSON format with the following JSON schema, and do not use markdown in the response:

{
    "contentType": "'Text' or 'AdaptiveCard' only",
    "content": "{The content of the response, may be plain text, or JSON based adaptive card}"
}`);

weatherAgent.activity(ActivityTypes.Message, async (context, state) => {
  const llmResponse = await agent.invoke(
    {
      messages: [sysMessage, new HumanMessage(context.activity.text!)],
    },
    {
      configurable: { thread_id: context.activity.conversation!.id },
    }
  );

  const llmResponseContent: WeatherForecastAgentResponse = JSON.parse(
    llmResponse.messages[llmResponse.messages.length - 1].content as string
  );

  if (llmResponseContent.contentType === "Text") {
    await context.sendActivity(llmResponseContent.content);
  } else if (llmResponseContent.contentType === "AdaptiveCard") {
    const response = MessageFactory.attachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: llmResponseContent.content,
    });
    await context.sendActivity(response);
  }
});
