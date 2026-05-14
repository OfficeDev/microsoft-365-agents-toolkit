// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Message Extension template artifact names matching the template repository folder names.
 * These must match the directory names inside template zip archives.
 */
export const MessageExtensionTemplateNames = {
  SearchBased: "message-extension-v2",
} as const;

/**
 * Create a standard scaffold function for message extension templates.
 */
function makeMEScaffoldFn(templateName: string) {
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
 * All Message Extension template descriptors.
 */
export const messageExtensionTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "me/search-based",
    name: "Search-based Message Extension",
    description: "A search-based message extension for Copilot",
    category: "message-extension",
    languages: ["typescript", "python", "csharp"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.SearchBased),
    displayOrder: 1,
    tags: ["teamsApp", "publishable"],
  },
];
