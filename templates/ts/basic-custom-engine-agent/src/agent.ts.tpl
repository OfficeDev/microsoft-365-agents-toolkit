import { ActivityTypes } from "@microsoft/agents-activity";
import { AgentApplication, MemoryStorage, TurnContext } from "@microsoft/agents-hosting";
import { AzureOpenAI } from "openai";
import config from "./config";

{{#useOpenAI}}
const client = new OpenAI({
  apiKey: config.openAIKey,
});
{{/useOpenAI}}
{{#useAzureOpenAI}}
const client = new AzureOpenAI({
  apiVersion: "2024-12-01-preview",
  apiKey: config.azureOpenAIKey,
  endpoint: config.azureOpenAIEndpoint,
  deployment: config.azureOpenAIDeploymentName,
});
{{/useAzureOpenAI}}
const systemPrompt = "You are an AI bot that can chat with users.";

// Define storage and application
const storage = new MemoryStorage();
export const agentApp = new AgentApplication({
  storage,
});

agentApp.conversationUpdate("membersAdded", async (context: TurnContext) => {
  await context.sendActivity(`Hi there! I'm an echo bot.`);
});

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
agentApp.activity(ActivityTypes.Message, async (context: TurnContext) => {
  // Echo back users request
  const result = await client.chat.completions.create({
    messages: [
      {
        role: "system",
        content: systemPrompt,
      },
      {
        role: "user",
        content: context.activity.text,
      },
    ],
    {{#useOpenAI}}
    model: config.openAIModelName
    {{/useOpenAI}}
    {{#useAzureOpenAI}}
    model: "",
    {{/useAzureOpenAI}}
  });
  let answer = "";
  for (const choice of result.choices) {
    answer += choice.message.content;
  }
  await context.sendActivity(answer);
});
