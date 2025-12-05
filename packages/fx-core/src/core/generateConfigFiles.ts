// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { FxError, Result, ok } from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { TOOLS } from "../common/globalVars";

export async function generateConfigFiles(
  appManifestFilePath: string,
  programmingLanguage: string,
  includePlayground: boolean,
  includeLocalDebug: boolean,
  includeRemoteDeploy: boolean
): Promise<Result<undefined, FxError>> {
  await TOOLS.ui?.showMessage(
    "info",
    `${includePlayground.toString()} ${includeLocalDebug.toString()} ${includeRemoteDeploy.toString()} ${programmingLanguage}`,
    false
  );
  return ok(undefined);
}
