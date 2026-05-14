/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach } from "mocha";
import { ok } from "neverthrow";
import { buildQuestionTree } from "../../../src/questions/treeBuilder";
import { QuestionNames } from "../../../src/questions/questionNames";
import { TemplateRegistry } from "../../../src/templates/registry";
import type { TemplateDescriptor, Language } from "../../../src/templates/types";

function makeDescriptor(
  overrides: Partial<TemplateDescriptor> & { id: string }
): TemplateDescriptor {
  return {
    name: overrides.id,
    category: "bot",
    languages: ["typescript"] as Language[],
    scaffoldFn: async () => ok({ projectPath: "/tmp" }),
    ...overrides,
  } as TemplateDescriptor;
}

describe("QuestionTreeBuilder", () => {
  let registry: TemplateRegistry;

  beforeEach(() => {
    registry = new TemplateRegistry();
  });

  it("should build an empty tree from empty registry", () => {
    const tree = buildQuestionTree(registry);
    expect(tree.data).to.have.property("type", "group");
    expect(tree.children).to.have.length(3); // projectType, projectName, destinationFolder
  });

  it("should include projectType, projectName, and destinationFolder at root", () => {
    registry.register(makeDescriptor({ id: "bot/echo", category: "bot" }));
    const tree = buildQuestionTree(registry);
    const childNames = tree.children!.map((c) => (c.data as any).name);
    expect(childNames).to.include(QuestionNames.projectType);
    expect(childNames).to.include(QuestionNames.projectName);
    expect(childNames).to.include(QuestionNames.destinationFolder);
  });

  it("should create category options from registered templates", () => {
    registry.register(makeDescriptor({ id: "bot/echo", category: "bot" }));
    registry.register(makeDescriptor({ id: "tab/basic", category: "tab" }));

    const tree = buildQuestionTree(registry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const options = (projectTypeNode.data as any).staticOptions;
    const optionIds = options.map((o: any) => o.id);
    expect(optionIds).to.include("bot");
    expect(optionIds).to.include("tab");
  });

  it("should order categories by predefined order", () => {
    registry.register(makeDescriptor({ id: "tab/basic", category: "tab" }));
    registry.register(makeDescriptor({ id: "da/basic", category: "declarative-agent" }));
    registry.register(makeDescriptor({ id: "bot/echo", category: "bot" }));

    const tree = buildQuestionTree(registry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const options = (projectTypeNode.data as any).staticOptions;
    const optionIds = options.map((o: any) => o.id);
    expect(optionIds.indexOf("declarative-agent")).to.be.lessThan(optionIds.indexOf("bot"));
    expect(optionIds.indexOf("bot")).to.be.lessThan(optionIds.indexOf("tab"));
  });

  it("should create template selector per category", () => {
    registry.register(makeDescriptor({ id: "bot/echo", category: "bot", name: "Echo Bot" }));
    registry.register(
      makeDescriptor({ id: "bot/workflow", category: "bot", name: "Workflow Bot" })
    );

    const tree = buildQuestionTree(registry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;

    // Should have one child node for the "bot" category
    const botCategoryNode = projectTypeNode.children!.find(
      (c) => (c as any).condition?.equals === "bot"
    );
    expect(botCategoryNode).to.exist;
    expect((botCategoryNode!.data as any).name).to.equal(QuestionNames.templateId);
    const templateOptions = (botCategoryNode!.data as any).staticOptions;
    expect(templateOptions).to.have.length(2);
  });

  it("should create language selector per template", () => {
    registry.register(
      makeDescriptor({
        id: "bot/echo",
        category: "bot",
        languages: ["typescript", "javascript"],
      })
    );

    const tree = buildQuestionTree(registry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const botCategoryNode = projectTypeNode.children!.find(
      (c) => (c as any).condition?.equals === "bot"
    )!;

    // Template selector children = one node per template (group)
    const templateGroupNode = botCategoryNode.children!.find(
      (c) => (c as any).condition?.equals === "bot/echo"
    );
    expect(templateGroupNode).to.exist;
    // Inside the template group, find the language question
    const langNode = templateGroupNode!.children!.find(
      (c) => (c.data as any).name === QuestionNames.language
    );
    expect(langNode).to.exist;
    expect((langNode!.data as any).staticOptions).to.have.length(2);
  });

  it("should include template-specific questions", () => {
    registry.register(
      makeDescriptor({
        id: "da/mcp-remote",
        category: "declarative-agent",
        languages: ["common"],
        questions: [
          {
            question: {
              type: "text",
              name: QuestionNames.mcpServerUrl,
              title: "MCP URL",
            },
          },
        ],
      })
    );

    const tree = buildQuestionTree(registry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const daCategoryNode = projectTypeNode.children!.find(
      (c) => (c as any).condition?.equals === "declarative-agent"
    )!;
    const mcpTemplateNode = daCategoryNode.children!.find(
      (c) => (c as any).condition?.equals === "da/mcp-remote"
    )!;

    const questionNames = mcpTemplateNode.children!.map((c) => (c.data as any).name);
    expect(questionNames).to.include(QuestionNames.language);
    expect(questionNames).to.include(QuestionNames.mcpServerUrl);
  });

  it("should filter out templates gated by disabled feature flags", () => {
    registry.register(makeDescriptor({ id: "da/basic", category: "declarative-agent" }));
    registry.register(
      makeDescriptor({
        id: "da/metaos",
        category: "declarative-agent",
        featureFlag: "DAMetaOS",
      })
    );

    // With no feature flag registry, flagged templates should be excluded
    const tree = buildQuestionTree(registry, undefined);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const daCategoryNode = projectTypeNode.children!.find(
      (c) => (c as any).condition?.equals === "declarative-agent"
    )!;
    const templateOptions = (daCategoryNode.data as any).staticOptions;
    expect(templateOptions).to.have.length(1);
    expect(templateOptions[0].id).to.equal("da/basic");
  });

  it("should include feature-flagged templates when flag is enabled", () => {
    registry.register(makeDescriptor({ id: "da/basic", category: "declarative-agent" }));
    registry.register(
      makeDescriptor({
        id: "da/metaos",
        category: "declarative-agent",
        featureFlag: "DAMetaOS",
      })
    );

    const mockFlagRegistry = {
      isEnabled: (flag: string) => flag === "DAMetaOS",
    } as any;

    const tree = buildQuestionTree(registry, mockFlagRegistry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const daCategoryNode = projectTypeNode.children!.find(
      (c) => (c as any).condition?.equals === "declarative-agent"
    )!;
    const templateOptions = (daCategoryNode.data as any).staticOptions;
    expect(templateOptions).to.have.length(2);
  });

  it("should sort templates by displayOrder within category", () => {
    registry.register(
      makeDescriptor({
        id: "bot/workflow",
        category: "bot",
        name: "Workflow",
        displayOrder: 2,
      })
    );
    registry.register(
      makeDescriptor({
        id: "bot/echo",
        category: "bot",
        name: "Echo",
        displayOrder: 1,
      })
    );

    const tree = buildQuestionTree(registry);
    const projectTypeNode = tree.children!.find(
      (c) => (c.data as any).name === QuestionNames.projectType
    )!;
    const botCategoryNode = projectTypeNode.children!.find(
      (c) => (c as any).condition?.equals === "bot"
    )!;
    const templateOptions = (botCategoryNode.data as any).staticOptions;
    expect(templateOptions[0].id).to.equal("bot/echo");
    expect(templateOptions[1].id).to.equal("bot/workflow");
  });
});
