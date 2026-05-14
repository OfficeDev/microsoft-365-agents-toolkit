/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { Command } from "commander";
import {
  TemplateRegistry,
  type TemplateDescriptor,
  type TemplateCategory,
  type QuestionSpec,
} from "@microsoft/teamsfx-core-next";
import {
  buildNewCommands,
  mapQuestionToOption,
  DEFAULT_SLUG_MAP,
} from "../../src/commands/factory";

function makeDescriptor(
  id: string,
  category: TemplateCategory,
  overrides: Partial<TemplateDescriptor> = {}
): TemplateDescriptor {
  return {
    id,
    name: id.replace("/", " "),
    category,
    languages: ["typescript", "javascript"],
    scaffoldFn: async () =>
      ({
        isOk: () => true,
        isErr: () => false,
        value: { projectPath: "/tmp", warnings: [] },
      }) as any,
    ...overrides,
  };
}

describe("Command Factory", () => {
  describe("buildNewCommands()", () => {
    it("should generate subcommands for each category with descriptors", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot"));
      registry.register(makeDescriptor("da/basic", "declarative-agent"));

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const subcmdNames = parent.commands.map((c) => c.name());
      expect(subcmdNames).to.include("bot");
      expect(subcmdNames).to.include("da");
    });

    it("should NOT generate subcommands for categories with no descriptors", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot"));

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const subcmdNames = parent.commands.map((c) => c.name());
      expect(subcmdNames).to.not.include("da");
      expect(subcmdNames).to.not.include("ai");
      expect(subcmdNames).to.not.include("tab");
    });

    it("should use configured category slug mapping", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("da/basic", "declarative-agent"));

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const subcmdNames = parent.commands.map((c) => c.name());
      expect(subcmdNames).to.include(DEFAULT_SLUG_MAP["declarative-agent"]);
    });

    it("should allow custom slug mapping", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot"));

      const parent = new Command("new");
      buildNewCommands(parent, registry, {
        slugMap: { bot: "bots" },
      });

      const subcmdNames = parent.commands.map((c) => c.name());
      expect(subcmdNames).to.include("bots");
    });

    it("should create template subcommands under category", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot"));
      registry.register(makeDescriptor("bot/workflow", "bot"));

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const botCmd = parent.commands.find((c) => c.name() === "bot");
      expect(botCmd).to.exist;
      const templateNames = botCmd!.commands.map((c) => c.name());
      expect(templateNames).to.include("echo");
      expect(templateNames).to.include("workflow");
    });

    it("should include --name (required) and --folder on every template", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot"));

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const botCmd = parent.commands.find((c) => c.name() === "bot");
      const echoCmd = botCmd!.commands.find((c) => c.name() === "echo");
      const optNames = echoCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--name");
      expect(optNames).to.include("--folder");
    });

    it("should include --language for multi-language templates", () => {
      const registry = new TemplateRegistry();
      registry.register(
        makeDescriptor("bot/echo", "bot", {
          languages: ["typescript", "javascript"],
        })
      );

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const botCmd = parent.commands.find((c) => c.name() === "bot");
      const echoCmd = botCmd!.commands.find((c) => c.name() === "echo");
      const optNames = echoCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--language");
    });

    it("should omit --language for single-language templates", () => {
      const registry = new TemplateRegistry();
      registry.register(
        makeDescriptor("da/basic", "declarative-agent", {
          languages: ["common"],
        })
      );

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const daCmd = parent.commands.find((c) => c.name() === "da");
      const basicCmd = daCmd!.commands.find((c) => c.name() === "basic");
      const optNames = basicCmd!.options.map((o) => o.long);
      expect(optNames).to.not.include("--language");
    });

    it("should add template-specific question options", () => {
      const questionSpec: QuestionSpec = {
        question: {
          name: "mcpServerUrl",
          type: "text" as const,
          title: "MCP server URL",
          cliName: "mcp-server-url",
          cliDescription: "URL of the MCP server",
        },
      };
      const registry = new TemplateRegistry();
      registry.register(
        makeDescriptor("da/mcp-remote", "declarative-agent", {
          questions: [questionSpec],
          languages: ["common"],
        })
      );

      const parent = new Command("new");
      buildNewCommands(parent, registry);

      const daCmd = parent.commands.find((c) => c.name() === "da");
      const mcpCmd = daCmd!.commands.find((c) => c.name() === "mcp-remote");
      const optNames = mcpCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--mcp-server-url");
    });

    it("extensibility: adding a new template to registry creates a new CLI command", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot"));

      // First build
      const parent1 = new Command("new");
      buildNewCommands(parent1, registry);
      const botCmd1 = parent1.commands.find((c) => c.name() === "bot");
      expect(botCmd1!.commands.length).to.equal(1);

      // Register a new template
      registry.register(makeDescriptor("bot/proactive", "bot"));

      // Second build
      const parent2 = new Command("new");
      buildNewCommands(parent2, registry);
      const botCmd2 = parent2.commands.find((c) => c.name() === "bot");
      expect(botCmd2!.commands.length).to.equal(2);
      expect(botCmd2!.commands.map((c) => c.name())).to.include("proactive");
    });
  });

  describe("mapQuestionToOption()", () => {
    it("should create a value option for text questions", () => {
      const spec: QuestionSpec = {
        question: {
          name: "apiSpecPath",
          type: "text" as const,
          title: "API spec path",
          cliName: "api-spec-path",
        },
      };
      const option = mapQuestionToOption(spec);
      expect(option).to.exist;
      expect(option!.long).to.equal("--api-spec-path");
    });

    it("should create a boolean flag for confirm questions", () => {
      const spec: QuestionSpec = {
        question: {
          name: "skipValidation",
          type: "confirm" as const,
          title: "Skip validation?",
          cliName: "skip-validation",
          isBoolean: true,
        },
      };
      const option = mapQuestionToOption(spec);
      expect(option).to.exist;
      expect(option!.long).to.equal("--skip-validation");
    });

    it("should return undefined for cliHidden questions", () => {
      const spec: QuestionSpec = {
        question: {
          name: "hidden",
          type: "text" as const,
          title: "Hidden",
          cliHidden: true,
        },
      };
      const option = mapQuestionToOption(spec);
      expect(option).to.be.undefined;
    });

    it("should use cliShortName as short flag", () => {
      const spec: QuestionSpec = {
        question: {
          name: "specPath",
          type: "text" as const,
          title: "Spec path",
          cliName: "spec-path",
          cliShortName: "s",
        },
      };
      const option = mapQuestionToOption(spec);
      expect(option).to.exist;
      expect(option!.short).to.equal("-s");
    });
  });
});
