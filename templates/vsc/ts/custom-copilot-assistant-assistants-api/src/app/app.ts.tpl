import { MemoryStorage, MessageFactory, TurnContext } from "botbuilder";

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
import { Application, AI, preview } from "@microsoft/teams-ai";

import config from "../config";

{{#useOpenAI}}
// See README.md to prepare your own OpenAI Assistant
if (!config.openAIKey || !config.openAIAssistantId) {
  throw new Error(
    "Missing OPENAI_API_KEY or OPENAI_ASSISTANT_ID. See README.md to prepare your own OpenAI Assistant."
  );
}
{{/useOpenAI}}
 {{#useAzureOpenAI}}
// See README.md to prepare your own Azure OpenAI Assistant
if (!config.azureOpenAIKey || !config.azureOpenAIAssistantId) {
  throw new Error(
    "Missing AZURE_OPENAI_API_KEY or AZURE_OPENAI_ASSISTANT_ID. See README.md to prepare your own Azure OpenAI Assistant."
  );
}
{{/useAzureOpenAI}}

import { resetMessage } from "./messages";
import { httpErrorAction, getCurrentWeather, getNickname } from "./actions";

// Create AI components
// Use OpenAI
const planner = new preview.AssistantsPlanner({
  {{#useOpenAI}}
  apiKey: config.openAIKey,
  assistant_id: config.openAIAssistantId,
  {{/useOpenAI}}
  {{#useAzureOpenAI}}
   apiKey: config.azureOpenAIKey,
   assistant_id: config.azureOpenAIAssistantId,
   endpoint: config.azureOpenAIEndpoint
  {{/useAzureOpenAI}}
});

// Define storage and application
const storage = new MemoryStorage();
const app = new Application({
  storage,
  ai: {
    planner,
    enable_feedback_loop: true,
  },
});

app.feedbackLoop(async (context, state, feedbackLoopData) => {
  //add custom feedback process logic here
  console.log("Your feedback is " + JSON.stringify(context.activity.value));
});

app.message("reset", resetMessage);

app.ai.action(AI.HttpErrorActionName, httpErrorAction);
app.ai.action("getCurrentWeather", getCurrentWeather);
app.ai.action("getNickname", getNickname);

export default app;
