import { QuestionNames } from "../questionNames";
import { TemplateNames } from "../templateName";

export const ceaNode = {
  condition: {
    equals: "custom-engine-agent-type",
  },
  data: {
    title: "question.customEngineAgent.title",
    name: QuestionNames.customEngineAgentType,
    type: "singleSelect",
    options: [
      {
        id: TemplateNames.BasicCustomEngineAgent,
        label: "question.customEngineAgent.option.basic.label",
        detail: "question.customEngineAgent.option.basic.detail",
        data: TemplateNames.BasicCustomEngineAgent,
      },
      {
        id: TemplateNames.WeatherAgent,
        label: "question.customEngineAgent.option.weather.label",
        detail: "question.customEngineAgent.option.weather.detail",
        data: TemplateNames.WeatherAgent,
      },
    ],
    placeholder: "question.placeholder.choose",
  },
  children: [
    {
      node: "llmServiceNode",
    },
  ],
};
