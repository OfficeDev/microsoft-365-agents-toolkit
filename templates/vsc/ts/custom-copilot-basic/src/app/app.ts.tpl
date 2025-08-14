import { App } from "@microsoft/teams.apps";
import { ChatPrompt } from "@microsoft/teams.ai";
import { LocalStorage } from "@microsoft/teams.common";
import { OpenAIChatModel } from "@microsoft/teams.openai";
import { MessageActivity } from '@microsoft/teams.api';
import * as fs from 'fs';
import * as path from 'path';
import config from "../config";

// Create storage for conversation history
const storage = new LocalStorage();

// Load instructions from file on initialization
function loadInstructions(): string {
  const instructionsFilePath = path.join(__dirname, "instructions.txt");
  return fs.readFileSync(instructionsFilePath, 'utf-8').trim();
}

// Load instructions once at startup
const instructions = loadInstructions();

// Create the app with storage
const app = new App({
  storage
});

// Handle application errors
app.event('error', async ({ error, send }) => {
  console.error('Error processing message:', error);
  await send("The agent encountered an error or bug.");
  await send("To continue to run this agent, please fix the agent source code.");
});

// Handle incoming messages
app.on('message', async ({ send, stream, activity }) => {
  //Get conversation history
  const conversationKey = `${activity.conversation.id}/${activity.from.id}`;
  const messages = storage.get(conversationKey) || [];

  const prompt = new ChatPrompt({
    messages,
    instructions,
    {{#useOpenAI}}
    model: new OpenAIChatModel({
      model: config.openAIModelName,
      apiKey: config.openAIKey
    })
    {{/useOpenAI}}
    {{#useAzureOpenAI}}
    model: new OpenAIChatModel({
      model: config.azureOpenAIDeploymentName,
      apiKey: config.azureOpenAIKey,
      endpoint: config.azureOpenAIEndpoint,
      apiVersion: "2024-10-21"
    })
    {{/useAzureOpenAI}}
  })

  if (activity.conversation.isGroup) {
    // If the conversation is a group chat, we need to send the final response
    // back to the group chat
    const response = await prompt.send(activity.text);
    const responseActivity = new MessageActivity(response.content).addAiGenerated().addFeedback();
    await send(responseActivity);
  } else {
      await prompt.send(activity.text, {
        onChunk: (chunk) => {
          stream.emit(chunk);
        },
      });
    // We wrap the final response with an AI Generated indicator
    stream.emit(new MessageActivity().addAiGenerated().addFeedback());
  }
  storage.set(conversationKey, messages);

});

app.on('message.submit.feedback', async ({ activity, log }) => {
  //add custom feedback process logic here
  console.log("Your feedback is " + JSON.stringify(activity.value));
  return {} as any;
})

export default app;