const { MemoryStorage, MessageFactory } = require("botbuilder");
const path = require("path");
const config = require("../config");
const customSayCommand = require("./customSayCommand");

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
const { AI, Application, ActionPlanner, OpenAIModel, PromptManager } = require("@microsoft/teams-ai");
const { MyDataSource } = require("./myDataSource");

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
const planner = new ActionPlanner({
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
const app = new Application({
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

module.exports = app;
