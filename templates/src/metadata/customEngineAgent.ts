// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const customEngineAgentTemplates: Template[] = [
  {
    id: "basic-custom-engine-agent-ts",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "typescript",
    description: "Build an AI agent with your own reasoning engine and custom logic",
  },
  {
    id: "basic-custom-engine-agent-js",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "javascript",
    description: "Build an AI agent with your own reasoning engine and custom logic",
  },
  {
    id: "basic-custom-engine-agent-python",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "python",
    description: "Build an AI agent with your own reasoning engine and custom logic",
  },
  {
    id: "weather-agent-ts",
    name: TemplateNames.WeatherAgent,
    language: "typescript",
    description: "AI agent that provides weather information using custom engine",
  },
  {
    id: "weather-agent-js",
    name: TemplateNames.WeatherAgent,
    language: "javascript",
    description: "AI agent that provides weather information using custom engine",
  },
  {
    id: "custom-copilot-weather-agent-csharp",
    name: TemplateNames.WeatherAgent,
    language: "csharp",
    description: "AI agent that provides weather information using custom engine",
  },
  {
    id: "custom-copilot-travel-agent-csharp",
    name: TemplateNames.TravelAgent,
    language: "csharp",
    description: "",
  },
];
