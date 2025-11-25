// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { getString } from "../ui/helper";
import { Template } from "./interface";

export const customEngineAgentTemplates: Template[] = [
  {
    id: "basic-custom-engine-agent-ts",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "typescript",
    displayName: getString("template.customEngineAgent.basic.label"),
    description: getString("template.customEngineAgent.basic.detail"),
  },
  {
    id: "basic-custom-engine-agent-js",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "javascript",
    displayName: getString("template.customEngineAgent.basic.label"),
    description: getString("template.customEngineAgent.basic.detail"),
  },
  {
    id: "basic-custom-engine-agent-python",
    name: TemplateNames.BasicCustomEngineAgent,
    language: "python",
    displayName: getString("template.customEngineAgent.basic.label"),
    description: getString("template.customEngineAgent.basic.detail"),
  },
  {
    id: "weather-agent-ts",
    name: TemplateNames.WeatherAgent,
    language: "typescript",
    displayName: getString("template.customEngineAgent.weather.label"),
    description: getString("template.customEngineAgent.weather.detail"),
  },
  {
    id: "weather-agent-js",
    name: TemplateNames.WeatherAgent,
    language: "javascript",
    displayName: getString("template.customEngineAgent.weather.label"),
    description: getString("template.customEngineAgent.weather.detail"),
  },
  {
    id: "custom-copilot-weather-agent-csharp",
    name: TemplateNames.WeatherAgent,
    language: "csharp",
    displayName: getString("template.customEngineAgent.weather.label"),
    description: getString("template.customEngineAgent.weather.detail"),
  },
  {
    id: "custom-copilot-travel-agent-csharp",
    name: TemplateNames.TravelAgent,
    language: "csharp",
    displayName: getString("template.customEngineAgent.travel.label"),
    description: getString("template.customEngineAgent.travel.detail"),
  },
];
