import { MemoryStorage, MessageFactory, TurnContext } from "botbuilder";
import * as path from "path";
import config from "../config";
import * as customSayCommand  from "./customSayCommand";

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
import { AI, Application, ActionPlanner, OpenAIModel, PromptManager, TurnState } from "@microsoft/teams-ai";
import { GraphDataSource } from "./graphDataSource";

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
const graphDataSource = new GraphDataSource("graph-ai-search");
planner.prompts.addDataSource(graphDataSource);

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<TurnState>({
  storage,
  ai: {
    planner,
    enable_feedback_loop: true,
  },
  authentication: {
    settings: {
      graph: {
        scopes: ["Files.Read.All"],
        msalConfig: {
          auth: {
            clientId: process.env.AAD_APP_CLIENT_ID,
            clientSecret: process.env.AAD_APP_CLIENT_SECRET,
            authority: `${process.env.AAD_APP_OAUTH_AUTHORITY_HOST}/${process.env.AAD_APP_TENANT_ID}`
          }
        },
        signInLink: `https://${process.env.BOT_DOMAIN}/auth-start.html`,
      }
    },
    autoSignIn: true,
  }
});
app.ai.action(AI.SayCommandActionName, customSayCommand.sayCommand(true));

app.authentication.get("graph").onUserSignInSuccess(async (context, state) => {
  // Successfully logged in
  await context.sendActivity("You are successfully logged in. You can send a new message to talk to the bot.");
});

app.authentication.get("graph").onUserSignInFailure(async (context, state, error) => {
  // Failed to login
  await context.sendActivity("Failed to login");
  await context.sendActivity(`Error message: ${error.message}`);
});

app.feedbackLoop(async (context, state, feedbackLoopData) => {
  //add custom feedback process logic here
  console.log("Your feedback is " + JSON.stringify(context.activity.value));
});

export default app;
