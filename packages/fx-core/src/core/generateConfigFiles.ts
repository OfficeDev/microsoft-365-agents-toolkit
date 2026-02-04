// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AppManifestUtils,
  FxError,
  Inputs,
  Result,
  TeamsManifest,
  TeamsManifestV1D24,
  ok,
} from "@microsoft/teamsfx-api";
import path from "path";
import "reflect-metadata";
import { TOOLS, createContext } from "../common/globalVars";
import { configGenerator } from "../component/generator/configFiles/configGenerator";
import { QuestionNames } from "../question/questionNames";

export async function generateConfigFiles(inputs: Inputs): Promise<Result<undefined, FxError>> {
  const appManifestFilePath = inputs["manifest-file"] as string;
  const projectPath = inputs[QuestionNames.ProjectPath] as string;
  const includePlayground = inputs["include-playground"];
  const includeLocalDebug = inputs["include-local"];
  const includeRemoteDeploy = inputs["include-remote"];
  const programmingLanguage = inputs["programming-language"] as string;

  const appManifest = await AppManifestUtils.readTeamsManifest(
    path.join(projectPath, appManifestFilePath)
  );
  const appName = appManifest.name.short;
  const features = detectAppFeatures(appManifest);
  const configComponents: { name: string; programmingLanguage: string }[] = [];

  if (includePlayground) {
    if (features.hasBot) {
      configComponents.push({ name: "playground", programmingLanguage });
    } else {
      await TOOLS.ui?.showMessage(
        "warn",
        `Playground is not supported for the current app manifest of ${appName}. Skipping Playground configuration file generation.`,
        false
      );
    }
  }

  if (includeLocalDebug) {
    configComponents.push({ name: "local", programmingLanguage });
  }

  if (includeRemoteDeploy) {
    configComponents.push({ name: "remote", programmingLanguage });
  }

  const context = createContext();
  await configGenerator.run(context, projectPath, configComponents, { ...features, appName });

  return ok(undefined);
}

function detectAppFeatures(manifest: TeamsManifest): Record<string, boolean> {
  const features: Record<string, boolean> = {};
  const manifestV1D24 = manifest as TeamsManifestV1D24.TeamsManifestV1D24;
  features["hasTab"] = manifest.staticTabs !== undefined && manifest.staticTabs.length > 0;
  features["hasBot"] = manifest.bots !== undefined && manifest.bots.length > 0;
  features["hasCopilot"] = manifestV1D24.copilotAgents !== undefined;
  features["hasMessageExtension"] =
    manifest.composeExtensions !== undefined && manifest.composeExtensions.length > 0;
  features["hasDeclarativeAgent"] =
    manifestV1D24.copilotAgents !== undefined &&
    manifestV1D24.copilotAgents.declarativeAgents !== undefined &&
    manifestV1D24.copilotAgents.declarativeAgents.length > 0;
  features["hasCustomEngineAgent"] =
    manifestV1D24.copilotAgents !== undefined &&
    manifestV1D24.copilotAgents.customEngineAgents !== undefined &&
    manifestV1D24.copilotAgents.customEngineAgents.length > 0;

  features["hasAzureBot"] = features["hasBot"] || features["hasMessageExtension"];
  features["supportCopilot"] = features["hasDeclarativeAgent"] || features["hasCustomEngineAgent"];
  return features;
}
