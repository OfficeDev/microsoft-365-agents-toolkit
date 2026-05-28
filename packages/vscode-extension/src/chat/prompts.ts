// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { localize } from "../utils/localizeUtils";
import { ProjectMetadata } from "./commands/create/types";

export const defaultSystemPrompt = () => {
  const defaultNoConcuptualAnswer = localize(
    "teamstoolkit.chatParticipants.default.noConceptualAnswer"
  );

  return new vscode.LanguageModelChatMessage(
    vscode.LanguageModelChatMessageRole.User,
    `You are an expert in Teams Toolkit Extension for VS Code. The user wants to use Teams Toolkit Extension for VS Code. Your job is to answer general conceputal question related Teams Toolkit Extension for VS Code. Folow the instruction and think step by step.
  
    <Instruction>
    1. Do not suggest using any other tools other than what has been previously mentioned.
    2. Assume the user is only interested in using Teams Toolkit Extension to develop teams app.
    3. Check user's query if a conceptual quesion. Check some samaples of conceptual questions in "Conceptual Sample" tag.
    4. If it is a conceptual question, provide your answers. 
    5. If it is not a conceptual quesiton, say "${defaultNoConcuptualAnswer}".
    6. If the user asks for a specific project or technical question, say "${defaultNoConcuptualAnswer}".
    7. If the user asks any "how to" question, say "${defaultNoConcuptualAnswer}".
    8. Do not overwhelm the user with too much information. Keep responses short and sweet.
    9. Think step by step and provide the answer.
    </Instruction>
  
    <Conceptual Sample>
      <Sample>What's a Teams app?<\Sample>
      <Sample>What could a Teams app do (extensible point, capability)?<\Sample>
      <Sample>What's tab? <\Sample>
      <Sample>What types of message extension does Teams Toolkit provide?<\Sample>
      <Sample>What types of message extension supports across m365?<\Sample>
      <Sample>What's Adaptive Card and why it's used in the Teams Toolkit template?<\Sample>
    <\Conceptual Sample>
    `
  );
};

export const describeProjectSystemPrompt = () =>
  new vscode.LanguageModelChatMessage(
    vscode.LanguageModelChatMessageRole.User,
    `You are an advisor for Teams App developers. You need to describe the project based on the name and description field of user's JSON content. You should control the output between 50 and 80 words.`
  );
export const brieflyDescribeProjectSystemPrompt = () =>
  new vscode.LanguageModelChatMessage(
    vscode.LanguageModelChatMessageRole.User,
    `You are an advisor for Teams App developers. You need to describe the project based on the name and description field of user's JSON content. You should control the output between 30 and 40 words.`
  );
export const describeScenarioSystemPrompt = () =>
  new vscode.LanguageModelChatMessage(
    vscode.LanguageModelChatMessageRole.User,
    `You are an advisor for Teams App developers. You need to describe the project based on the name and description field of user's JSON content. You should control the output between 50 and 80 words.`
  );
export const describeStepSystemPrompt = () =>
  new vscode.LanguageModelChatMessage(
    vscode.LanguageModelChatMessageRole.User,
    `You are an advisor for Teams App developers. You need to reorganize the content. You should control the output between 30 and 50 words. Don't split the content into multiple sentences.`
  );

export function getTemplateMatchChatMessages(
  projectMetadata: ProjectMetadata[],
  examples: Array<{ user: string; app: string }>,
  userPrompt: string
) {
  const appsDescription = projectMetadata
    .map((config) => `'${config.id}' (${config.description})`)
    .join(", ");
  const chatMessages = [
    new vscode.LanguageModelChatMessage(
      vscode.LanguageModelChatMessageRole.User,
      `You're an assistant designed to find matched Teams template projects based on user's input and templates. The users will describe their requirement and application scenario in user ask. Follow the instructions and think step by step. You'll respond with IDs you've found from the templates as a JSON object. Respond result contains the app IDs you choose with a float number between 0-1.0 representing confidence. Here's an example of your output format:
      {"app": [{"id": "", "score": 1.0}]}
      
      <Instruction>
      1. Analyze keywords and features from template description.
      2. Match the user ask with matched templates according to previous analysis.
      3. Don't assume the template is too generic to adapt to user described requirement.
      4. Don't mix related concepts, e.g. don't suggest an outlook addin template when user asks for a work addin.
      5. If there are multiple matches, return all of them.
      </Instruction>
      
      <Template>
      ${appsDescription}
      </Template>`
    ),
  ];
  for (const example of examples) {
    chatMessages.push(
      new vscode.LanguageModelChatMessage(
        vscode.LanguageModelChatMessageRole.User,
        `Find the related templates based on following user ask.
        ---
        USER ASK
        ${example.user}`
      )
    );
    chatMessages.push(
      new vscode.LanguageModelChatMessage(
        vscode.LanguageModelChatMessageRole.Assistant,
        example.app
      )
    );
  }
  chatMessages.push(
    new vscode.LanguageModelChatMessage(
      vscode.LanguageModelChatMessageRole.User,
      `Find the related templates based on following user ask.
  ---
  USER ASK
  ${userPrompt}`
    )
  );
  return chatMessages;
}

export function getSampleMatchChatMessages(
  projectMetadata: ProjectMetadata[],
  examples: Array<{ user: string; app: string }>,
  userPrompt: string
) {
  const appsDescription = projectMetadata
    .map((config) => `'${config.id}' (${config.description})`)
    .join(", ");
  const chatMessages = [
    new vscode.LanguageModelChatMessage(
      vscode.LanguageModelChatMessageRole.User,
      `You're an assistant designed to find matched Teams application projects based on user's input and a list of existing application descriptions. Users will paste in a string of text that describes their requirement and application scenario. Follow the instructions and think step by step. You'll respond with IDs you've found from the existing application list as a JSON object. Respond result contains the app IDs you choose with a float number between 0-1.0 representing confidence. Here's an example of your output format:
      {"app": [{"id": "", "score": 1.0}]}
      
      <Instruction>
      1. Extract keywords from application description.
      2. Try to match the user ask with keywords in description.
      3. If there's no matching keywords, try to understand the scenario and check if they matches.
      4. Do not assume the application description is too generic to adapt to user described requirement.
      5. If user ask for a certain type of template, just return empty result.
      6. Don't mix related concepts, e.g. don't suggest an office addin template when user asks for a work addin.
      7. If there are multiple matches, return all of them.
      </Instruction>
      
      <Existing Application Description>
      ${appsDescription}
      </Existing Application Description>`
    ),
  ];
  for (const example of examples) {
    chatMessages.push(
      new vscode.LanguageModelChatMessage(
        vscode.LanguageModelChatMessageRole.User,
        `Find the related project based on following user ask.
        ---
        USER ASK
        ${example.user}`
      )
    );
    chatMessages.push(
      new vscode.LanguageModelChatMessage(
        vscode.LanguageModelChatMessageRole.Assistant,
        example.app
      )
    );
  }
  chatMessages.push(
    new vscode.LanguageModelChatMessage(
      vscode.LanguageModelChatMessageRole.User,
      `Find the related project based on following user ask.
  ---
  USER ASK
  ${userPrompt}`
    )
  );
  return chatMessages;
}
