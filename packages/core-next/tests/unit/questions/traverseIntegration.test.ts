/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { ok } from "neverthrow";
import { createMockContext } from "../testHelper";
import { buildQuestionTree } from "../../../src/questions/treeBuilder";
import { traverseQuestionTree } from "../../../src/questions/traverse";
import { QuestionNames } from "../../../src/questions/questionNames";
import { TemplateRegistry } from "../../../src/templates/registry";
import type { TemplateDescriptor, Language } from "../../../src/templates/types";
import type { InputResult } from "../../../src/api/qm/ui";
import { Platform } from "../../../src/api/constants";
import type { Inputs } from "../../../src/api/types";

/**
 * Integration test: builds a question tree from a test registry, traverses it
 * with a mock UserInteraction, and verifies the resulting inputs are correct.
 */
describe("Question tree integration – build + traverse", () => {
  let sandbox: sinon.SinonSandbox;
  let ctx: ReturnType<typeof createMockContext>;
  let registry: TemplateRegistry;
  let tmpDir: string;

  const stubScaffold = sinon.stub().resolves(ok({ warnings: [] }));

  const descriptors: TemplateDescriptor[] = [
    {
      id: "bot/basic",
      name: "Basic Bot",
      category: "bot",
      languages: ["typescript", "javascript"] as Language[],
      scaffoldFn: stubScaffold,
      displayOrder: 1,
    },
    {
      id: "bot/ai",
      name: "AI Bot",
      category: "bot",
      languages: ["typescript", "python"] as Language[],
      scaffoldFn: stubScaffold,
      displayOrder: 2,
      questions: [
        {
          question: {
            type: "singleSelect",
            name: QuestionNames.llmProvider,
            title: "LLM Provider",
            staticOptions: [
              { id: "azure-openai", label: "Azure OpenAI" },
              { id: "openai", label: "OpenAI" },
            ],
          } as any,
        },
      ],
    },
    {
      id: "tab/react",
      name: "React Tab",
      category: "tab",
      languages: ["typescript"] as Language[],
      scaffoldFn: stubScaffold,
      displayOrder: 1,
    },
  ];

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    ctx = createMockContext();
    registry = new TemplateRegistry();
    for (const d of descriptors) {
      registry.register(d);
    }
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "traverse-integ-"));
  });

  afterEach(() => {
    sandbox.restore();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it("should collect all required inputs for a bot template via tree traversal", async () => {
    const tree = buildQuestionTree(registry);
    const inputs: Inputs = { platform: Platform.VSCode };

    const selectStub = ctx.ui.selectOption as sinon.SinonStub;
    const inputStub = ctx.ui.inputText as sinon.SinonStub;
    const folderStub = ctx.ui.selectFolder as sinon.SinonStub;

    // Queue answers in order: projectType → templateId → language → projectName → folder
    // projectType: "bot"
    selectStub.onCall(0).resolves(ok({ type: "success", result: "bot" } as InputResult<string>));
    // templateId: "bot/basic"
    selectStub
      .onCall(1)
      .resolves(ok({ type: "success", result: "bot/basic" } as InputResult<string>));
    // language: "typescript"
    selectStub
      .onCall(2)
      .resolves(ok({ type: "success", result: "typescript" } as InputResult<string>));
    // projectName
    inputStub.onCall(0).resolves(ok({ type: "success", result: "my-bot" } as InputResult<string>));
    // destination folder
    folderStub.onCall(0).resolves(ok({ type: "success", result: tmpDir } as InputResult<string>));

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;

    // Verify all required inputs were collected
    expect(inputs[QuestionNames.projectType]).to.equal("bot");
    expect(inputs[QuestionNames.templateId]).to.equal("bot/basic");
    expect(inputs[QuestionNames.language]).to.equal("typescript");
    expect(inputs[QuestionNames.projectName]).to.equal("my-bot");
    expect(inputs[QuestionNames.destinationFolder]).to.equal(tmpDir);
  });

  it("should collect extra questions for templates with custom questions", async () => {
    const tree = buildQuestionTree(registry);
    const inputs: Inputs = { platform: Platform.VSCode };

    const selectStub = ctx.ui.selectOption as sinon.SinonStub;
    const inputStub = ctx.ui.inputText as sinon.SinonStub;
    const folderStub = ctx.ui.selectFolder as sinon.SinonStub;

    // projectType: "bot"
    selectStub.onCall(0).resolves(ok({ type: "success", result: "bot" } as InputResult<string>));
    // templateId: "bot/ai" (has LLM question)
    selectStub.onCall(1).resolves(ok({ type: "success", result: "bot/ai" } as InputResult<string>));
    // language: "typescript"
    selectStub
      .onCall(2)
      .resolves(ok({ type: "success", result: "typescript" } as InputResult<string>));
    // llmProvider extra question
    selectStub
      .onCall(3)
      .resolves(ok({ type: "success", result: "azure-openai" } as InputResult<string>));
    // projectName
    inputStub.onCall(0).resolves(ok({ type: "success", result: "ai-bot" } as InputResult<string>));
    // destination folder
    folderStub.onCall(0).resolves(ok({ type: "success", result: tmpDir } as InputResult<string>));

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;

    expect(inputs[QuestionNames.templateId]).to.equal("bot/ai");
    expect(inputs[QuestionNames.llmProvider]).to.equal("azure-openai");
    expect(inputs[QuestionNames.projectName]).to.equal("ai-bot");
  });

  it("should work with pre-filled projectType to skip first question", async () => {
    const tree = buildQuestionTree(registry);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.projectType]: "tab",
    };

    const selectStub = ctx.ui.selectOption as sinon.SinonStub;
    const inputStub = ctx.ui.inputText as sinon.SinonStub;
    const folderStub = ctx.ui.selectFolder as sinon.SinonStub;

    // templateId: only tab/react is available, may auto-skip if skipSingleOption
    selectStub
      .onCall(0)
      .resolves(ok({ type: "success", result: "tab/react" } as InputResult<string>));
    // language: only typescript, may auto-skip
    selectStub
      .onCall(1)
      .resolves(ok({ type: "success", result: "typescript" } as InputResult<string>));
    // projectName
    inputStub.onCall(0).resolves(ok({ type: "success", result: "my-tab" } as InputResult<string>));
    // folder
    folderStub.onCall(0).resolves(ok({ type: "success", result: tmpDir } as InputResult<string>));

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;

    // projectType should remain pre-filled
    expect(inputs[QuestionNames.projectType]).to.equal("tab");
    expect(inputs[QuestionNames.projectName]).to.equal("my-tab");
  });
});
