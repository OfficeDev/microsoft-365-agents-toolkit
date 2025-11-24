// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const customEngineAgentTemplates: Template[] = [
  {
    id: "basic-custom-engine-agent-ts",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "typescript",
    displayName: "Basic Custom Engine Agent",
    description: "Build an AI agent with your own reasoning engine and custom logic",
  },
  {
    id: "basic-custom-engine-agent-js",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "javascript",
    displayName: "Basic Custom Engine Agent",
    description: "Build an AI agent with your own reasoning engine and custom logic",
  },
  {
    id: "basic-custom-engine-agent-python",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "python",
    displayName: "Basic Custom Engine Agent",
    description: "Build an AI agent with your own reasoning engine and custom logic",
  },
  {
    id: "weather-agent-ts",
    name: TemplateNames.WeatherAgent,
    language: "typescript",
    displayName: "Weather Agent",
    description: "AI agent that provides weather information using custom engine",
  },
  {
    id: "weather-agent-js",
    name: TemplateNames.WeatherAgent,
    language: "javascript",
    displayName: "Weather Agent",
    description: "AI agent that provides weather information using custom engine",
  },
  {
    id: "custom-copilot-weather-agent-csharp",
    name: TemplateNames.WeatherAgent,
    language: "csharp",
    displayName: "Weather Agent",
    description: "AI agent that provides weather information using custom engine",
  },
  {
    id: "custom-copilot-travel-agent-csharp",
    name: TemplateNames.TravelAgent,
    language: "csharp",
    displayName: "Travel Agent",
    description:
      "AI agent that provides comprehensive travel assistance by answering travel-related questions, helping users understand company travel policies, and finding flights and hotels that comply with organizational guidelines.",
  },
];
