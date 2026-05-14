// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";
import {
  graphConnectorNameQuestion,
  graphConnectorConnectionIdQuestion,
} from "../../questions/commonQuestions";

/**
 * Connector template artifact names matching the template repository folder names.
 */
export const ConnectorTemplateNames = {
  Graph: "graph-connector",
} as const;

/**
 * Create a standard scaffold function for connector templates.
 */
function makeConnectorScaffoldFn(templateName: string) {
  return async (ctx: AtkContext, opts: TemplateActionOptions) => {
    const replaceMap: Record<string, string> = {
      ...getTemplateReplaceMap({
        appName: opts.projectName,
        ...opts,
      }),
    };

    if (typeof opts.graphConnectorName === "string") {
      replaceMap.gcName = opts.graphConnectorName;
    }
    if (typeof opts.graphConnectorConnectionId === "string") {
      replaceMap.gcConnectionId = opts.graphConnectorConnectionId;
    }

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
 * All Connector template descriptors.
 */
export const connectorTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "connector/graph",
    name: "Graph Connector",
    description: "A Microsoft Graph connector for ingesting external data",
    category: "connector",
    languages: ["typescript"],
    scaffoldFn: makeConnectorScaffoldFn(ConnectorTemplateNames.Graph),
    displayOrder: 1,
    // ARM deploys flaky Microsoft.ApplicationInsights.AzureWebSites site extension +
    // Key Vault + RBAC; not reliably testable in CI without real connector secrets.
    testable: false,
    questions: [graphConnectorNameQuestion(), graphConnectorConnectionIdQuestion()],
  },
];
