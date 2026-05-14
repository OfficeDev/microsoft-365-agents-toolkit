/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { daTemplateDescriptors } from "../../../src/templates/descriptors/declarativeAgent";
import { botTemplateDescriptors } from "../../../src/templates/descriptors/bot";
import { tabTemplateDescriptors } from "../../../src/templates/descriptors/tab";
import { aiAgentTemplateDescriptors } from "../../../src/templates/descriptors/aiAgent";
import { engineAgentTemplateDescriptors } from "../../../src/templates/descriptors/engineAgent";
import { connectorTemplateDescriptors } from "../../../src/templates/descriptors/connector";
import { messageExtensionTemplateDescriptors } from "../../../src/templates/descriptors/messageExtension";
import { openApiTemplateDescriptors } from "../../../src/templates/descriptors/openApi";
import { foundryTemplateDescriptors } from "../../../src/templates/descriptors/foundry";
import type { TemplateDescriptor } from "../../../src/templates/types";
import { QuestionNames } from "../../../src/questions/questionNames";

/**
 * Verify common invariants for a descriptor array.
 */
function describeDescriptorArray(name: string, descriptors: TemplateDescriptor[]) {
  describe(name, () => {
    it("should have unique IDs", () => {
      const ids = descriptors.map((d) => d.id);
      expect(new Set(ids).size).to.equal(ids.length, `Duplicate IDs found in ${name}`);
    });

    it("should have non-empty names", () => {
      for (const d of descriptors) {
        expect(d.name).to.be.a("string").and.not.be.empty;
      }
    });

    it("should have at least one supported language", () => {
      for (const d of descriptors) {
        expect(d.languages).to.be.an("array").with.length.greaterThan(0);
      }
    });

    it("should have a scaffoldFn function", () => {
      for (const d of descriptors) {
        expect(d.scaffoldFn).to.be.a("function");
      }
    });

    it("should have valid displayOrder values", () => {
      for (const d of descriptors) {
        if (d.displayOrder !== undefined) {
          expect(d.displayOrder).to.be.a("number").and.greaterThan(0);
        }
      }
    });

    it("should have displayOrder values in ascending sequence within each category", () => {
      const byCategory = new Map<string, number[]>();
      for (const d of descriptors) {
        if (d.displayOrder !== undefined) {
          const list = byCategory.get(d.category) ?? [];
          list.push(d.displayOrder);
          byCategory.set(d.category, list);
        }
      }
      for (const [_cat, orders] of byCategory) {
        for (let i = 1; i < orders.length; i++) {
          expect(orders[i]).to.be.greaterThanOrEqual(orders[i - 1]);
        }
      }
    });

    it("should have IDs matching the category/variant convention", () => {
      for (const d of descriptors) {
        expect(d.id).to.match(/^[a-z-]+\/[a-z0-9-]+$/);
      }
    });
  });
}

describe("Template Descriptors", () => {
  describeDescriptorArray("DA Descriptors", daTemplateDescriptors);
  describeDescriptorArray("Bot Descriptors", botTemplateDescriptors);
  describeDescriptorArray("Tab Descriptors", tabTemplateDescriptors);
  describeDescriptorArray("AI Agent Descriptors", aiAgentTemplateDescriptors);
  describeDescriptorArray("Engine Agent Descriptors", engineAgentTemplateDescriptors);
  describeDescriptorArray("Connector Descriptors", connectorTemplateDescriptors);
  describeDescriptorArray("Message Extension Descriptors", messageExtensionTemplateDescriptors);
  describeDescriptorArray("OpenAPI Descriptors", openApiTemplateDescriptors);
  describeDescriptorArray("Foundry Descriptors", foundryTemplateDescriptors);

  describe("All descriptors combined", () => {
    const all = [
      ...daTemplateDescriptors,
      ...botTemplateDescriptors,
      ...tabTemplateDescriptors,
      ...aiAgentTemplateDescriptors,
      ...engineAgentTemplateDescriptors,
      ...connectorTemplateDescriptors,
      ...messageExtensionTemplateDescriptors,
      ...openApiTemplateDescriptors,
      ...foundryTemplateDescriptors,
    ];

    it("should have globally unique IDs", () => {
      const ids = all.map((d) => d.id);
      const unique = new Set(ids);
      expect(unique.size).to.equal(
        ids.length,
        `Duplicate IDs: ${ids.filter((id, i) => ids.indexOf(id) !== i)}`
      );
    });

    it("should have correct categories", () => {
      for (const d of daTemplateDescriptors) {
        expect(d.category).to.equal("declarative-agent");
      }
      for (const d of botTemplateDescriptors) {
        expect(d.category).to.equal("bot");
      }
      for (const d of tabTemplateDescriptors) {
        expect(d.category).to.equal("tab");
      }
      for (const d of aiAgentTemplateDescriptors) {
        expect(d.category).to.equal("ai-agent");
      }
      for (const d of foundryTemplateDescriptors) {
        expect(d.category).to.equal("ai-agent");
      }
      for (const d of engineAgentTemplateDescriptors) {
        expect(d.category).to.equal("custom-engine-agent");
      }
      for (const d of connectorTemplateDescriptors) {
        expect(d.category).to.equal("connector");
      }
      for (const d of messageExtensionTemplateDescriptors) {
        expect(d.category).to.equal("message-extension");
      }
    });
  });

  describe("Question metadata", () => {
    it("AI agent descriptors should have LLM questions", () => {
      for (const d of aiAgentTemplateDescriptors) {
        expect(d.questions).to.be.an("array").with.length.greaterThan(0);
        const questionNames = d.questions!.map((q) => q.question.name);
        expect(questionNames).to.include(QuestionNames.llmProvider);
      }
    });

    it("connector/graph should have graph connector questions", () => {
      const gc = connectorTemplateDescriptors.find((d) => d.id === "connector/graph")!;
      expect(gc.questions).to.be.an("array").with.length(2);
      const questionNames = gc.questions!.map((q) => q.question.name);
      expect(questionNames).to.include(QuestionNames.graphConnectorName);
      expect(questionNames).to.include(QuestionNames.graphConnectorConnectionId);
    });

    it("da/mcp-remote should have mcpServerUrl question", () => {
      const mcp = daTemplateDescriptors.find((d) => d.id === "da/mcp-remote")!;
      expect(mcp.questions).to.be.an("array").with.length(1);
      expect(mcp.questions![0].question.name).to.equal(QuestionNames.mcpServerUrl);
    });

    it("bot descriptors should have no extra questions", () => {
      for (const d of botTemplateDescriptors) {
        expect(d.questions).to.be.undefined;
      }
    });

    it("tab descriptors should have no extra questions", () => {
      for (const d of tabTemplateDescriptors) {
        expect(d.questions).to.be.undefined;
      }
    });

    it("foundry descriptors should have foundry questions", () => {
      for (const d of foundryTemplateDescriptors) {
        expect(d.questions).to.be.an("array").with.length(2);
        const questionNames = d.questions!.map((q) => q.question.name);
        expect(questionNames).to.include(QuestionNames.foundryEndpoint);
        expect(questionNames).to.include(QuestionNames.foundryAgentId);
      }
    });
  });

  describe("Feature flags", () => {
    it("da/metaos should be gated by DAMetaOS flag", () => {
      const metaos = daTemplateDescriptors.find((d) => d.id === "da/metaos")!;
      expect(metaos.featureFlag).to.equal("DAMetaOS");
    });

    it("da/metaos-upgrade should be gated by DAMetaOS flag", () => {
      const upgrade = daTemplateDescriptors.find((d) => d.id === "da/metaos-upgrade")!;
      expect(upgrade.featureFlag).to.equal("DAMetaOS");
      expect(upgrade.questions).to.be.an("array").with.length(1);
      expect(upgrade.questions![0].question.name).to.equal(QuestionNames.officeAddinFolder);
    });

    it("non-flagged descriptors should not have featureFlag", () => {
      const unflagged = [
        ...botTemplateDescriptors,
        ...tabTemplateDescriptors,
        ...aiAgentTemplateDescriptors,
        ...engineAgentTemplateDescriptors,
        ...connectorTemplateDescriptors,
        ...messageExtensionTemplateDescriptors,
        ...foundryTemplateDescriptors,
      ];
      for (const d of unflagged) {
        expect(d.featureFlag).to.be.undefined;
      }
    });
  });

  describe("Descriptor counts", () => {
    it("should have 12 DA descriptors", () => {
      expect(daTemplateDescriptors).to.have.length(12);
    });

    it("should have 1 bot descriptor", () => {
      expect(botTemplateDescriptors).to.have.length(1);
    });

    it("should have 1 tab descriptor", () => {
      expect(tabTemplateDescriptors).to.have.length(1);
    });

    it("should have 3 AI agent descriptors", () => {
      expect(aiAgentTemplateDescriptors).to.have.length(3);
    });

    it("should have 3 engine agent descriptors", () => {
      expect(engineAgentTemplateDescriptors).to.have.length(3);
    });

    it("should have 1 connector descriptor", () => {
      expect(connectorTemplateDescriptors).to.have.length(1);
    });

    it("should have 1 message extension descriptor", () => {
      expect(messageExtensionTemplateDescriptors).to.have.length(1);
    });

    it("should have 3 OpenAPI descriptors", () => {
      expect(openApiTemplateDescriptors).to.have.length(3);
    });

    it("should have 1 foundry descriptor", () => {
      expect(foundryTemplateDescriptors).to.have.length(1);
    });

    it("should have 26 total descriptors", () => {
      const total =
        daTemplateDescriptors.length +
        botTemplateDescriptors.length +
        tabTemplateDescriptors.length +
        aiAgentTemplateDescriptors.length +
        engineAgentTemplateDescriptors.length +
        connectorTemplateDescriptors.length +
        messageExtensionTemplateDescriptors.length +
        openApiTemplateDescriptors.length +
        foundryTemplateDescriptors.length;
      expect(total).to.equal(26);
    });
  });
});
