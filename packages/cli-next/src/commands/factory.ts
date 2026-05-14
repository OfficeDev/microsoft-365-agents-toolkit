// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command, Option } from "commander";
import type {
  TemplateRegistry,
  TemplateDescriptor,
  TemplateCategory,
  QuestionSpec,
  UserInputQuestion,
} from "@microsoft/teamsfx-core-next";
import { wrapHandlerWithContext } from "../handler";
import { createProjectAction } from "../actions/createProject";

/**
 * Default mapping from TemplateCategory to CLI slug.
 * Consumers can override via the `slugMap` parameter.
 */
export const DEFAULT_SLUG_MAP: Record<TemplateCategory, string> = {
  "declarative-agent": "da",
  "custom-engine-agent": "cea",
  "ai-agent": "ai",
  "message-extension": "me",
  bot: "bot",
  tab: "tab",
  connector: "connector",
  "office-addin": "addin",
};

/**
 * Human-readable descriptions for each category slug.
 */
const CATEGORY_DESCRIPTIONS: Record<TemplateCategory, string> = {
  "declarative-agent": "Declarative Agent templates",
  "custom-engine-agent": "Custom Engine Agent templates",
  "ai-agent": "AI Agent templates",
  "message-extension": "Message Extension templates",
  bot: "Bot templates",
  tab: "Tab templates",
  connector: "Connector templates",
  "office-addin": "Office Add-in templates",
};

/**
 * All recognized template categories.
 */
const ALL_CATEGORIES: TemplateCategory[] = [
  "declarative-agent",
  "custom-engine-agent",
  "ai-agent",
  "message-extension",
  "bot",
  "tab",
  "connector",
  "office-addin",
];

export interface BuildNewCommandsOptions {
  slugMap?: Partial<Record<TemplateCategory, string>>;
}

/**
 * Build `atk new <category> <template>` subcommands from the TemplateRegistry.
 *
 * For each category that has at least one registered template, creates a
 * category subcommand (e.g. `atk new da`). Under each category, creates
 * a template subcommand (e.g. `atk new da basic`) with options derived
 * from the descriptor's questions and languages.
 */
export function buildNewCommands(
  parent: Command,
  registry: TemplateRegistry,
  opts?: BuildNewCommandsOptions
): void {
  const slugMap = { ...DEFAULT_SLUG_MAP, ...(opts?.slugMap ?? {}) };

  for (const category of ALL_CATEGORIES) {
    const descriptors = registry.listByCategory(category);
    if (descriptors.length === 0) continue;

    const slug = slugMap[category];
    const catCmd = parent.command(slug).description(CATEGORY_DESCRIPTIONS[category]);

    for (const descriptor of descriptors) {
      buildTemplateCommand(catCmd, descriptor);
    }
  }
}

/**
 * Build a single template subcommand under a category.
 */
function buildTemplateCommand(parent: Command, descriptor: TemplateDescriptor): void {
  // Template slug is the part after the "/" in the ID (e.g. "da/basic" → "basic")
  const slug = descriptor.id.includes("/")
    ? descriptor.id.split("/").slice(1).join("-")
    : descriptor.id;

  const cmd = parent.command(slug).description(descriptor.description ?? descriptor.name);

  // Common options
  cmd.requiredOption("-n, --name <name>", "Project name");
  cmd.option("-f, --folder <folder>", "Target folder", ".");

  // Language option — only if multiple real languages
  const realLanguages = descriptor.languages.filter((l) => l !== "common");
  if (realLanguages.length > 1) {
    cmd.option(
      "-l, --language <language>",
      `Programming language (${realLanguages.join(", ")})`,
      realLanguages[0]
    );
  }

  // Template-specific questions → Commander options
  if (descriptor.questions) {
    for (const spec of descriptor.questions) {
      const option = mapQuestionToOption(spec);
      if (option) {
        cmd.addOption(option);
      }
    }
  }

  // Wire handler
  cmd.action(
    wrapHandlerWithContext(`new ${slug}`, async (ctx, cmdOpts) => {
      const language =
        (cmdOpts.language as string) ??
        (realLanguages.length === 1 ? realLanguages[0] : undefined) ??
        (descriptor.languages.includes("common") ? "common" : descriptor.languages[0]);

      await createProjectAction(ctx, {
        templateId: descriptor.id,
        projectName: cmdOpts.name as string,
        language,
        destinationPath: (cmdOpts.folder as string) ?? ".",
        options: cmdOpts as Record<string, unknown>,
      });
    })
  );
}

/**
 * Convert a QuestionSpec into a Commander Option.
 * Returns undefined if the question should be skipped (cliHidden, etc).
 */
export function mapQuestionToOption(spec: QuestionSpec): Option | undefined {
  const q = spec.question as UserInputQuestion;
  if (!q || q.cliHidden) return undefined;

  const name = q.cliName ?? q.name;
  const shortFlag = q.cliShortName ? `-${q.cliShortName}, ` : "";
  const description = q.cliDescription ?? (typeof q.title === "string" ? q.title : name);

  if (q.isBoolean || q.type === "confirm") {
    return new Option(`${shortFlag}--${name}`, description);
  }

  return new Option(`${shortFlag}--${name} <value>`, description);
}
