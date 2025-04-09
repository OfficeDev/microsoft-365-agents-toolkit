// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Context,
  err,
  FxError,
  InputsWithProjectPath,
  Result,
  TeamsAppManifest,
  UserError,
  ok,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import fs from "fs-extra";
import { DriverContext } from "../component/driver/interface/commonArgs";
import { getLocalizedString } from "./localizeUtils";
import { pathUtils } from "../component/utils/pathUtils";
import { metadataUtil } from "../component/utils/metadataUtil";
import { resolve } from "../component/configManager/lifecycle";
import { DriverDefinition } from "../component/configManager/interface";
import * as path from "path";
import AdmZip from "adm-zip";
import { Constants } from "../component/driver/teamsApp/constants";
import { envUtil } from "../component/utils/envUtil";

export async function waitSeconds(second: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, second * 1000));
}

export function generateDriverContext(ctx: Context, inputs: InputsWithProjectPath): DriverContext {
  return {
    azureAccountProvider: ctx.tokenProvider!.azureAccountProvider,
    m365TokenProvider: ctx.tokenProvider!.m365TokenProvider,
    ui: ctx.userInteraction,
    progressBar: undefined,
    logProvider: ctx.logProvider,
    telemetryReporter: ctx.telemetryReporter,
    projectPath: inputs.projectPath,
    platform: inputs.platform,
  };
}

export async function isJsonSpecFile(filePath: string): Promise<boolean> {
  const specPath = filePath.toLowerCase();
  if (specPath.endsWith(".yaml") || specPath.endsWith(".yml")) {
    return false;
  } else if (specPath.endsWith(".json")) {
    return true;
  }
  const isRemoteFile = specPath.startsWith("http:") || specPath.startsWith("https:");

  try {
    const fileContent = isRemoteFile
      ? (await axios.get(specPath)).data
      : await fs.readFile(specPath, "utf-8");
    JSON.parse(fileContent);
    return true;
  } catch (error) {
    return false;
  }
}
