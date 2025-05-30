import { MemoryStorage, MessageFactory, TurnContext } from "botbuilder";
import * as path from "path";
import config from "../config";
import * as customSayCommand  from "./customSayCommand";

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
import { AI, Application, ActionPlanner, OpenAIModel, PromptManager, TurnState } from "@microsoft/teams-ai";
import { MyDataSource } from "./myDataSource";

// Create AI components
const model = new OpenAIModel({
  {{#useOpenAI}}
  apiKey: config.openAIKey,
  defaultModel: config.openAIModelName,
  {{/useOpenAI}}
  {{#useAzureOpenAI}}
  azureApiKey: config.azureOpenAIKey,
  azureDefaultDeployment: config.azureOpenAIDeploymentName,
  azureEndpoint: config.azureOpenAIEndpoint,
  {{/useAzureOpenAI}}

  useSystemMessages: true,
  logRequests: true,
});
const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, "../prompts"),
});
const planner = new ActionPlanner<TurnState>({
  model,
  prompts,
  defaultPrompt: "chat",
});

// Register your data source with planner
const myDataSource = new MyDataSource("my-ai-search");
myDataSource.init();
planner.prompts.addDataSource(myDataSource);

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<TurnState>({
  storage,
  ai: {
    planner,
    enable_feedback_loop: true,
  },
});
app.ai.action(AI.SayCommandActionName, customSayCommand.sayCommand(true));

app.feedbackLoop(async (context, state, feedbackLoopData) => {
  //add custom feedback process logic here
  console.log("Your feedback is " + JSON.stringify(context.activity.value));
});

export default app;
