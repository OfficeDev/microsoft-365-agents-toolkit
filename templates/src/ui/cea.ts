// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { QuestionNames } from "../questionNames";
import { TemplateNames } from "../templateNames";

export const ceaNode = {
  condition: {
    equals: "custom-engine-agent-type",
  },
  data: {
    title: "template.customEngineAgent.title",
    name: QuestionNames.customEngineAgentType,
    type: "singleSelect",
    options: [
      {
        id: TemplateNames.BasicCustomEngineAgent,
        label: "template.customEngineAgent.basic.label",
        detail: "template.customEngineAgent.basic.detail",
        data: TemplateNames.BasicCustomEngineAgent,
      },
      {
        id: TemplateNames.WeatherAgent,
        label: "template.customEngineAgent.weather.label",
        detail: "template.customEngineAgent.weather.detail",
        data: TemplateNames.WeatherAgent,
      },
    ],
    placeholder: "template.customEngineAgent.placeholder",
  },
  children: [
    {
      node: "llmServiceNode",
    },
  ],
};
