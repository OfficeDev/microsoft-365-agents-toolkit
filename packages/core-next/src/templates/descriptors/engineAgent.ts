// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Custom Engine Agent template artifact names matching the template repository folder names.
 */
export const EngineAgentTemplateNames = {
  Basic: "basic-custom-engine-agent",
  FunctionCalling: "weather-agent",
  TeamsCollaborator: "teams-collaborator-agent",
} as const;

/**
 * csharp fallback zips use different folder names for some templates.
 * Map (base template name, language) → actual zip folder name.
 */
const csharpTemplateNameOverrides: Record<string, string> = {
  [EngineAgentTemplateNames.FunctionCalling]: "custom-copilot-weather-agent",
};

/**
 * Create a standard scaffold function for Custom Engine Agent templates.
 */
function makeEngineAgentScaffoldFn(templateName: string) {
  return async (ctx: AtkContext, opts: TemplateActionOptions) => {
    const replaceMap = getTemplateReplaceMap({
      appName: opts.projectName,
      ...opts,
    });

    const effectiveName =
      opts.language === "csharp"
        ? csharpTemplateNameOverrides[templateName] ?? templateName
        : templateName;

    const tplInfo: TemplateInfo = {
      templateName: effectiveName,
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
 * All Custom Engine Agent template descriptors.
 */
export const engineAgentTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "engine-agent/basic",
    name: "Basic Agent",
    description: "A basic custom engine agent",
    category: "custom-engine-agent",
    languages: ["typescript", "javascript", "python"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.Basic),
    displayOrder: 1,
    tags: ["teamsApp", "publishable", "bot"],
  },
  {
    id: "engine-agent/function-calling",
    name: "Agent with Function Calling",
    description: "A custom engine agent with function calling capabilities",
    category: "custom-engine-agent",
    languages: ["typescript", "javascript", "python", "csharp"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.FunctionCalling),
    displayOrder: 2,
    tags: ["teamsApp", "publishable", "bot"],
  },
  {
    id: "engine-agent/teams-collaborator",
    name: "Teams Collaborator Agent",
    description: "A custom engine agent for Teams collaboration",
    category: "custom-engine-agent",
    languages: ["typescript", "csharp"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.TeamsCollaborator),
    displayOrder: 4,
    tags: ["teamsApp", "publishable", "bot"],
  },
];
