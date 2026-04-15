// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Composite runtime nodes for the "Add Action" flow.
 * These encapsulate complex branching logic (feature flags, dynamic options)
 * that cannot be expressed in static JSON.
 *
 * Referenced by constructNode.ts via lazy require() to avoid circular dependencies.
 */

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { FeatureFlags, featureFlagManager } from "../../common/featureFlags";
import { getLocalizedString } from "../../common/localizeUtils";
import { apiOperationQuestion, apiSpecLocationQuestion } from "../create";
import { QuestionNames } from "../questionNames";
import { inputOrSearchAPISpecNode } from "./commonNodes";

/**
 * Composite node: API spec branch for Add Action.
 * Encapsulates KiotaNPM feature flag branching internally.
 * When KiotaNPM is enabled: uses inputOrSearchAPISpecNode (type select → URL/file/search)
 * When KiotaNPM is disabled: uses legacy apiSpecLocation + apiOperation flow
 */
export function addActionApiSpecNode(): IQTreeNode {
  return {
    data: { type: "group", name: "add-action-api-spec" },
    children: [
      {
        ...inputOrSearchAPISpecNode(),
        condition: () => featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration),
      },
      {
        data: apiSpecLocationQuestion(),
        condition: () => !featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration),
      },
      {
        data: apiOperationQuestion(true, true),
        condition: () => !featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration),
      },
    ],
  };
}

/**
 * Composite node: MCP for Declarative Agent.
 * Server URL → tools file path → auth type selection.
 */
export function mcpForDANode(): IQTreeNode {
  return {
    data: {
      name: QuestionNames.MCPForDAServerUrl,
      title: getLocalizedString("core.createProjectQuestion.mcpForDa.ServerUrl.title"),
      type: "text",
      placeholder: getLocalizedString(
        "core.createProjectQuestion.mcpForDa.ServerUrl.placeholder"
      ),
    },
    children: [
      {
        data: {
          name: QuestionNames.MCPToolsFilePath,
          title: getLocalizedString("core.MCPForDA.toolsFilePath.title"),
          type: "text",
          placeholder: getLocalizedString("core.MCPForDA.toolsFilePath.placeholder"),
        },
        children: [
          {
            data: {
              type: "singleSelect",
              name: QuestionNames.MCPForDAAuthType,
              title: getLocalizedString(
                "core.createProjectQuestion.mcpForDa.AuthType.title"
              ),
              staticOptions: [
                {
                  id: "oauth",
                  label: getLocalizedString(
                    "core.createProjectQuestion.mcpForDa.Auth.OAuth"
                  ),
                },
                {
                  id: "entraSSO",
                  label: getLocalizedString(
                    "core.createProjectQuestion.mcpForDa.Auth.EntraSSO"
                  ),
                },
              ],
              default: "oauth",
            },
          },
        ],
      },
    ],
  };
}
