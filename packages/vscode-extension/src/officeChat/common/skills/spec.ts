// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LanguageModelChatMessage } from "vscode";
import { SampleData } from "../samples/sampleData";
import { deepClone } from "../utils";

export class Spec {
  public userInput: string;
  public taskSummary: string;
  public sections: string[];
  public inspires: string[];
  public resources: string[];
  public appendix: {
    host: string;
    codeSnippet: string;
    codeExplanation: string;
    codeTaskBreakdown: string[];
    codeSample: string;
    apiDeclarationsReference: Map<string, SampleData>;
    isCustomFunction: boolean;
    telemetryData: {
      requestId: string;
      isHarmful: boolean;
      relatedSampleName: string[];
      chatMessages: LanguageModelChatMessage[];
      responseChatMessages: LanguageModelChatMessage[];
      properties: { [key: string]: string };
      measurements: { [key: string]: number };
    };
    complexity: number;
    shouldContinue: boolean;
  };

  constructor(userInput: string, requestId?: string) {
    this.userInput = userInput;
    this.taskSummary = "";
    this.sections = [];
    this.inspires = [];
    this.resources = [];
    this.appendix = {
      host: "",
      codeSnippet: "",
      codeExplanation: "",
      codeTaskBreakdown: [],
      codeSample: "",
      apiDeclarationsReference: new Map<string, SampleData>(),
      isCustomFunction: false,
      telemetryData: {
        requestId: requestId ? requestId : "",
        isHarmful: false,
        relatedSampleName: [],
        chatMessages: [],
        responseChatMessages: [],
        properties: {},
        measurements: {},
      },
      complexity: 0,
      shouldContinue: false,
    };
  }

  public clone(other: Spec): Spec {
    this.userInput = other.userInput;
    this.taskSummary = other.taskSummary;
    this.sections = other.sections;
    this.inspires = other.inspires;
    this.resources = other.resources;
    this.appendix = deepClone(other.appendix);
    return this;
  }
}
