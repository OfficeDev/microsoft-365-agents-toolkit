// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const customEngineAgentTemplates: Template[] = [
  {
    id: "basic-custom-engine-agent-ts",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "typescript",
    description: "",
  },
  {
    id: "basic-custom-engine-agent-js",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "javascript",
    description: "",
  },
  {
    id: "weather-agent-ts",
    name: TemplateNames.WeatherAgent,
    language: "typescript",
    description: "",
  },
  {
    id: "weather-agent-js",
    name: TemplateNames.WeatherAgent,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-weather-agent-csharp",
    name: TemplateNames.WeatherAgent,
    language: "csharp",
    description: "",
  },
];
