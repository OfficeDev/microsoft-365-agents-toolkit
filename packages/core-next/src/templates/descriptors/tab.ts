// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Tab template artifact names matching the template repository folder names.
 * These must match the directory names inside template zip archives.
 */
export const TabTemplateNames = {
  Basic: "basic-tab",
} as const;

/**
 * Create a standard scaffold function for tab templates.
 */
function makeTabScaffoldFn(templateName: string) {
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
 * All Tab template descriptors.
 */
export const tabTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "tab/basic",
    name: "Basic Tab",
    description: "A simple tab app without SSO",
    category: "tab",
    languages: ["typescript", "csharp"],
    scaffoldFn: makeTabScaffoldFn(TabTemplateNames.Basic),
    displayOrder: 1,
    tags: ["teamsApp", "publishable"],
  },
];
