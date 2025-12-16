const { preview } = require("@microsoft/teams-ai");

{{#useOpenAI}}
const openAIKey = process.argv[2];
if (!openAIKey) {
  throw new Error("Missing input OpenAI Key");
}
{{/useOpenAI}}
{{#useAzureOpenAI}}
const azureOpenAIKey = process.argv[2];
{{#azureOpenAIEndpoint}}
const azureOpenAIEndpoint="{{{azureOpenAIEndpoint}}}";
{{/azureOpenAIEndpoint}}
{{^azureOpenAIEndpoint}}
const azureOpenAIEndpoint="";
{{/azureOpenAIEndpoint}}
{{#azureOpenAIDeploymentName}}
const azureOpenAIDeploymentName="{{{azureOpenAIDeploymentName}}}";
{{/azureOpenAIDeploymentName}}
{{^azureOpenAIDeploymentName}}
const azureOpenAIDeploymentName="";
{{/azureOpenAIDeploymentName}}
if (!azureOpenAIKey || !azureOpenAIDeploymentName || !azureOpenAIEndpoint) {
  throw new Error("Missing input Azure OpenAI Key, Deployment Name or Endpoint");
}
{{/useAzureOpenAI}}

// Create new Assistant
(async () => {
{{#useOpenAI}}
  const assistant = await preview.AssistantsPlanner.createAssistant(openAIKey, {
{{/useOpenAI}}
{{#useAzureOpenAI}}
  const assistant = await preview.AssistantsPlanner.createAssistant(azureOpenAIKey, {
{{/useAzureOpenAI}}
    name: "Assistant",
    instructions: [
      "You are an intelligent agent that can",
      "- write and run code to answer math questions",
      "- use the provided functions to answer questions",
    ].join("\n"),
    tools: [
      {
        type: "code_interpreter",
      },
      {
        type: "function",
        function: {
          name: "getCurrentWeather",
          description: "Get the weather in location",
          parameters: {
            type: "object",
            properties: {
              location: {
                type: "string",
                description: "The city and state e.g. San Francisco, CA",
              },
              unit: {
                type: "string",
                enum: ["c", "f"],
              },
            },
            required: ["location"],
          },
        },
      },
      {
        type: "function",
        function: {
          name: "getNickname",
          description: "Get the nickname of a city",
          parameters: {
            type: "object",
            properties: {
              location: {
                type: "string",
                description: "The city and state e.g. San Francisco, CA",
              },
            },
            required: ["location"],
          },
        },
      },
    ],
{{#useOpenAI}}
    model: "gpt-3.5-turbo",
  });
{{/useOpenAI}}
{{#useAzureOpenAI}}
    model: azureOpenAIDeploymentName,
  },
  azureOpenAIEndpoint);
{{/useAzureOpenAI}}

  console.log(`Created a new assistant with an ID of: ${assistant.id}`);
})();
