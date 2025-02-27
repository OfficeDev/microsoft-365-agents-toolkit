// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author yuqzho@microsoft.com, KennethBWSong, Ning Tang
 */

import { ProjectType } from "@microsoft/m365-spec-parser";
import {
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
import { QuestionNames } from "../../../question";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";
import { defaultDeclarativeAgentActionId, defaultDeclarativeAgentManifestFileName } from "./const";
import {
  copyKiotaFolder,
  generateAdaptiveCardInPluginManifestForKiota,
  parseAndUpdatePluginManifestForKiota,
} from "./helper";

export function isKiotaIntegrated(inputs: Inputs): boolean {
  return (
    featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
    inputs[QuestionNames.DeclarativeAgentManifestPath]
  );
}

export async function getAuthDataFromKiota(
  context: Context,
  inputs: Inputs
): Promise<
  { authName: string; authType: "apiKey" | "oauth2"; registrationId: string }[] | undefined
> {
  // For Kiota integration, we need to get auth info here
  if (isKiotaIntegrated(inputs)) {
    const pluginManifestPath = inputs[QuestionNames.DeclarativeAgentManifestPath] as string;
    return await parseAndUpdatePluginManifestForKiota(pluginManifestPath, false);
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
  await fs.copyFile(inputs[QuestionNames.DeclarativeAgentManifestPath], pluginManifestPath);

  // 2.1 Need to update the plugin manifest file
  await parseAndUpdatePluginManifestForKiota(pluginManifestPath, true);

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
    defaultDeclarativeAgentManifestFileName,
    defaultDeclarativeAgentActionId,
    pluginManifestPath
  );
  if (addActionResult.isErr()) {
    return err(addActionResult.error);
  }

  // 5. Update plugin manifest to add ac info (optional)
  await generateAdaptiveCardInPluginManifestForKiota(pluginManifestPath, openapiSpecPath, context);

  // 5. Copy .kiota folder
  await copyKiotaFolder(inputs[QuestionNames.DeclarativeAgentManifestPath], destinationPath);
  return ok({ warnings: undefined });
}
