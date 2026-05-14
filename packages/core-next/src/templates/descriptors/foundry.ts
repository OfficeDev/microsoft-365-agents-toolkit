// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";
import { foundryEndpointQuestion, foundryAgentIdQuestion } from "../../questions/commonQuestions";

/**
 * Foundry Agent template artifact names matching the template repository folder names.
 */
export const FoundryTemplateNames = {
  FoundryAgent: "foundry-agent-to-m365",
} as const;

/**
 * Questions specific to Foundry Agent templates.
 */
const foundryQuestions = [foundryEndpointQuestion(), foundryAgentIdQuestion()];

/**
 * Create a standard scaffold function for Foundry Agent templates.
 */
function makeFoundryScaffoldFn(templateName: string) {
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
 * All Foundry Agent template descriptors.
 */
export const foundryTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "ai-agent/foundry-to-m365",
    name: "Foundry Agent",
    description: "An Microsoft 365 Agent that connects to Microsoft AI Foundry Agent",
    category: "ai-agent",
    languages: ["typescript"],
    scaffoldFn: makeFoundryScaffoldFn(FoundryTemplateNames.FoundryAgent),
    displayOrder: 4,
    questions: foundryQuestions,
    // Lifecycle E2E requires a Foundry endpoint + Azure App Service.
    // Enable once CI test tenant has a provisioned Foundry agent.
    testable: false,
  },
];
