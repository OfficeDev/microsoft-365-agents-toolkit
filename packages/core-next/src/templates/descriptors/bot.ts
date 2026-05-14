// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Bot template artifact names matching the template repository folder names.
 */
export const BotTemplateNames = {
  Echo: "default-bot",
} as const;

/**
 * Create a standard scaffold function for bot templates.
 * All bot templates share the same scaffold pipeline with no extra variables.
 */
function makeBotScaffoldFn(templateName: string) {
  return async (ctx: AtkContext, opts: TemplateActionOptions) => {
    const replaceMap = getTemplateReplaceMap({
      appName: opts.projectName,
      ...opts,
    });

    const tplInfo: TemplateInfo = {
      templateName,
      language: convertToLangKey(opts.language),
      replaceMap,
    };

    const result = await scaffoldTemplates(ctx, [tplInfo], opts.destinationPath);
    return result.map((files) => ({
      projectPath: opts.destinationPath,
      warnings: files.length === 0 ? ["No files were scaffolded"] : undefined,
    }));
  };
}

/**
 * All Bot template descriptors.
 */
export const botTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "bot/echo",
    name: "Echo Bot",
    description: "A simple bot that echoes user messages",
    category: "bot",
    languages: ["typescript", "javascript", "python", "csharp"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.Echo),
    displayOrder: 1,
    tags: ["teamsApp", "publishable", "bot"],
  },
];
