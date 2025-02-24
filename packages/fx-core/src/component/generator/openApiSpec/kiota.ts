// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author yuqzho@microsoft.com, Ning Tang
 */

import { ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import {
  AuthInfo,
  Context,
  FxError,
  GeneratorResult,
  Inputs,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { err, ok, Result } from "neverthrow";
import path from "path";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../../question";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";
import {
  defaultDeclarativeCopilotActionId,
  defaultDeclarativeCopilotManifestFileName,
} from "./const";
import { copyKiotaFolder, getParserOptions, listOperations } from "./helper";

export function isKiotaIntegrated(inputs: Inputs): boolean {
  return (
    featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
    inputs[QuestionNames.ApiPluginManifestPath]
  );
}

export async function getAuthDataFromKiota(
  context: Context,
  inputs: Inputs
): Promise<AuthInfo[] | undefined> {
  // For Kiota integration, we need to get auth info here
  if (isKiotaIntegrated(inputs)) {
    const operationsResult = await listOperations(
      context,
      inputs[QuestionNames.ApiSpecLocation],
      inputs
    );
    if (operationsResult.isErr()) {
      const errorMsg = getLocalizedString("error.kiota.FailedToGenerateAuthActions");
      void context.userInteraction.showMessage("warn", errorMsg, false);
      context.logProvider.warning(errorMsg);
    } else {
      const operations = operationsResult.value;
      const authApi = operations.filter((api) => !!api.data.authName);
      if (authApi.length > 0) {
        return authApi.map((api) => api.data);
      }
    }
  }
  return undefined;
}

export async function kiotaPostProcess(
  context: Context,
  inputs: Inputs,
  destinationPath: string,
  openapiSpecPath: string,
  pluginManifestPath: string,
  manifestPath: string,
  templateType: ProjectType,
  isDeclarativeAgent: boolean
): Promise<Result<GeneratorResult, FxError>> {
  // For Kiota integration scenario, we need to:
  // 1. Copy openapi spec file
  await fs.copyFile(inputs[QuestionNames.ApiSpecLocation].trim(), openapiSpecPath);

  // 2. Copy plugin manifest file
  await fs.copyFile(inputs[QuestionNames.ApiPluginManifestPath], pluginManifestPath);

  // 3. Update teams app manifest
  const manifest: TeamsAppManifest = await fs.readJSON(manifestPath);
  const apiPluginRelativePath = path.relative(manifestPath, pluginManifestPath);
  manifest.copilotAgents = manifest.copilotAgents || {};
  manifest.copilotAgents.plugins = [
    {
      file: apiPluginRelativePath,
      id: "plugin_1",
    },
  ];

  // 4. add action in da manifest
  const addActionResult = await copilotGptManifestUtils.updateDeclarativeAgentManifest(
    manifestPath,
    defaultDeclarativeCopilotManifestFileName,
    defaultDeclarativeCopilotActionId,
    pluginManifestPath
  );
  if (addActionResult.isErr()) {
    return err(addActionResult.error);
  }

  // 5. Update plugin manifest to add auth info (optional)
  try {
    const specParser = new SpecParser(
      openapiSpecPath,
      getParserOptions(templateType, isDeclarativeAgent)
    );
    const operation = (await specParser.list()).APIs.filter((value) => value.isValid).map(
      (value) => value.api
    );
    await specParser.generateAdaptiveCardInPlugin(pluginManifestPath, operation, undefined);
  } catch (error) {
    // create ac error, should not block the whole process
    const errorMsg = getLocalizedString("error.kiota.FailedToCreateAdaptiveCard");
    void context.userInteraction.showMessage("warn", errorMsg, false);
    context.logProvider.warning(errorMsg);
  }

  // 5. Copy .kiota folder
  await copyKiotaFolder(inputs[QuestionNames.ApiPluginManifestPath], destinationPath);
  return ok({ warnings: undefined });
}
