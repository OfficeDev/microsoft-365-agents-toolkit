/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import {
  projectNameQuestion,
  destinationFolderQuestion,
  languageQuestion,
  apiSpecPathQuestion,
  llmProviderQuestion,
  azureOpenAiKeyQuestion,
  openAiKeyQuestion,
  graphConnectorNameQuestion,
  graphConnectorConnectionIdQuestion,
  mcpServerUrlQuestion,
} from "../../../src/questions/commonQuestions";
import { QuestionNames } from "../../../src/questions/questionNames";
import { Platform } from "../../../src/api/constants";
import type { Inputs } from "../../../src/api/types";

describe("QuestionNames", () => {
  it("should have unique values", () => {
    const values = Object.values(QuestionNames);
    const unique = new Set(values);
    expect(unique.size).to.equal(values.length);
  });
});

describe("Common Question Factories", () => {
  describe("projectNameQuestion", () => {
    it("should create a text question with validation", () => {
      const q = projectNameQuestion();
      expect(q.type).to.equal("text");
      expect(q.name).to.equal(QuestionNames.projectName);
      expect(q.validation).to.have.property("maxLength", 30);
      expect(q.validation).to.have.property("pattern");
    });
  });

  describe("destinationFolderQuestion", () => {
    it("should create a folder question", () => {
      const q = destinationFolderQuestion();
      expect(q.type).to.equal("folder");
      expect(q.name).to.equal(QuestionNames.destinationFolder);
    });

    it("should default to './' on CLI", () => {
      const q = destinationFolderQuestion();
      const inputs: Inputs = { platform: Platform.CLI };
      const result = (q.default as (inputs: Inputs) => string)(inputs);
      expect(result).to.equal("./");
    });

    it("should default to home dir on VS Code", () => {
      const q = destinationFolderQuestion();
      const inputs: Inputs = { platform: Platform.VSCode };
      const result = (q.default as (inputs: Inputs) => string)(inputs);
      expect(result).to.include("TeamsApps");
    });
  });

  describe("languageQuestion", () => {
    it("should create a single-select with mapped labels", () => {
      const q = languageQuestion(["typescript", "javascript"]);
      expect(q.type).to.equal("singleSelect");
      expect(q.name).to.equal(QuestionNames.language);
      expect(q.staticOptions).to.have.length(2);
      expect((q.staticOptions[0] as any).label).to.equal("TypeScript");
      expect((q.staticOptions[1] as any).label).to.equal("JavaScript");
    });

    it("should skip when single option", () => {
      const q = languageQuestion(["common"]);
      expect(q.skipSingleOption).to.be.true;
    });
  });

  describe("apiSpecPathQuestion", () => {
    it("should create a singleFileOrText spec", () => {
      const spec = apiSpecPathQuestion();
      expect(spec.question.type).to.equal("singleFileOrText");
      expect(spec.question.name).to.equal(QuestionNames.apiSpecPath);
    });
  });

  describe("llmProviderQuestion", () => {
    it("should create a select with azure-openai and openai", () => {
      const spec = llmProviderQuestion();
      expect(spec.question.type).to.equal("singleSelect");
      expect(spec.question.name).to.equal(QuestionNames.llmProvider);
      const opts = (spec.question as any).staticOptions;
      expect(opts.map((o: any) => o.id)).to.deep.equal(["azure-openai", "openai"]);
    });
  });

  describe("azureOpenAiKeyQuestion", () => {
    it("should be conditional on llmProvider = azure-openai", () => {
      const spec = azureOpenAiKeyQuestion();
      expect(spec.dependsOn).to.equal(QuestionNames.llmProvider);
      expect(spec.condition!({ platform: Platform.CLI, llmProvider: "azure-openai" })).to.be.true;
      expect(spec.condition!({ platform: Platform.CLI, llmProvider: "openai" })).to.be.false;
    });
  });

  describe("openAiKeyQuestion", () => {
    it("should be conditional on llmProvider = openai", () => {
      const spec = openAiKeyQuestion();
      expect(spec.dependsOn).to.equal(QuestionNames.llmProvider);
      expect(spec.condition!({ platform: Platform.CLI, llmProvider: "openai" })).to.be.true;
      expect(spec.condition!({ platform: Platform.CLI, llmProvider: "azure-openai" })).to.be.false;
    });
  });

  describe("graphConnectorNameQuestion", () => {
    it("should create a text question with default", () => {
      const spec = graphConnectorNameQuestion();
      expect(spec.question.type).to.equal("text");
      expect(spec.question.name).to.equal(QuestionNames.graphConnectorName);
    });
  });

  describe("graphConnectorConnectionIdQuestion", () => {
    it("should depend on graphConnectorName", () => {
      const spec = graphConnectorConnectionIdQuestion();
      expect(spec.dependsOn).to.equal(QuestionNames.graphConnectorName);
      expect((spec.question as any).validation).to.have.property("pattern");
    });
  });

  describe("mcpServerUrlQuestion", () => {
    it("should create a text question", () => {
      const spec = mcpServerUrlQuestion();
      expect(spec.question.type).to.equal("text");
      expect(spec.question.name).to.equal(QuestionNames.mcpServerUrl);
    });
  });
});
