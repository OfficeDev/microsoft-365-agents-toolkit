// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AppManifestUtils,
  FxError,
  Inputs,
  Result,
  TeamsManifest,
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
  const configComponents: { name: string; programmingLanguage: string }[] = [];

  if (includePlayground) {
    if (isPlaygroundSupported(appManifest)) {
      configComponents.push({ name: "playground", programmingLanguage });
    } else {
      await TOOLS.ui?.showMessage(
        "warn",
        `Playground is not supported for the current app manifest of ${appName}. Skipping Playground configuration file generation.`,
        false
      );
    }
  }

  const context = createContext();
  await configGenerator.run(context, projectPath, configComponents);

  return ok(undefined);
}

function isPlaygroundSupported(manifest: TeamsManifest): boolean {
  if (manifest.bots && manifest.bots.length > 0) {
    return true;
  }
  return false;
}
