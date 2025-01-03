// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import isUUID from "validator/lib/isUUID";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { getUuid } from "../../../../common/stringUtils";
import {
  FileNotFoundError,
  JSONSyntaxError,
  MissingEnvironmentVariablesError,
} from "../../../../error/common";
import { expandEnvironmentVariable, getEnvironmentVariables } from "../../../utils/common";
import { DriverContext } from "../../interface/commonArgs";
import { MissingFieldInManifestUserError } from "../error/invalidFieldInManifestError";
import { AADManifest } from "../interface/AADManifest";
import { UpdateAadAppOutput } from "../interface/updateAadAppOutput";
import { logMessageKeys } from "../utility/constants";
import { AadManifestHelper } from "./aadManifestHelper";
import { AADApplication } from "../interface/AADApplication";

const actionName = "aadApp/update"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/aadapp-update";

export async function buildAadManifest(
  context: DriverContext,
  manifestPath: string,
  outputFilePath: string,
  state?: UpdateAadAppOutput
): Promise<AADManifest | AADApplication> {
  const manifestAbsolutePath = getAbsolutePath(manifestPath, context.projectPath);
  if (!(await fs.pathExists(manifestAbsolutePath))) {
    throw new FileNotFoundError(actionName, manifestAbsolutePath, helpLink);
  }
  const manifest = await loadManifest(manifestAbsolutePath, state);
  const warningMessage = AadManifestHelper.validateManifest(manifest);
  if (warningMessage) {
    warningMessage.split("\n").forEach((warning) => {
      context.logProvider?.warning(warning);
    });
  }

  if (!manifest.id || !isUUID(manifest.id)) {
    throw new MissingFieldInManifestUserError(actionName, "id", helpLink);
  }

  // Output actual manifest to project folder first for better troubleshooting experience
  const outputFileAbsolutePath = getAbsolutePath(outputFilePath, context.projectPath);
  await fs.ensureDir(path.dirname(outputFileAbsolutePath));
  await fs.writeFile(outputFileAbsolutePath, JSON.stringify(manifest, null, 4), "utf8");
  context.logProvider?.info(
    getLocalizedString(logMessageKeys.outputAadAppManifest, outputFileAbsolutePath)
  );

  return manifest;
}

function getAbsolutePath(relativeOrAbsolutePath: string, projectPath: string) {
  return path.isAbsolute(relativeOrAbsolutePath)
    ? relativeOrAbsolutePath
    : path.join(projectPath, relativeOrAbsolutePath);
}

async function loadManifest(
  manifestPath: string,
  state?: UpdateAadAppOutput
): Promise<AADManifest | AADApplication> {
  let generatedNewPermissionId = false;
  try {
    const manifestTemplate = await fs.readFile(manifestPath, "utf8");
    const permissionIdPlaceholderRegex = /\${{ *AAD_APP_ACCESS_AS_USER_PERMISSION_ID *}}/;

    // generate a new permission id if there's no one in env and manifest needs it
    if (!process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID) {
      const matches = permissionIdPlaceholderRegex.exec(manifestTemplate);
      if (matches) {
        const permissionId = getUuid();
        process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID = permissionId;
        if (state) {
          state.AAD_APP_ACCESS_AS_USER_PERMISSION_ID = permissionId;
        }
        generatedNewPermissionId = true;
      }
    }

    const manifestString = expandEnvironmentVariable(manifestTemplate);
    const unresolvedEnvironmentVariable = getEnvironmentVariables(manifestString);
    if (unresolvedEnvironmentVariable && unresolvedEnvironmentVariable.length > 0) {
      const error = new MissingEnvironmentVariablesError(
        "aadAppUpdate",
        unresolvedEnvironmentVariable.join(", "),
        manifestPath,
        helpLink
      );
      throw error;
    }
    let manifest: AADManifest | AADApplication;
    try {
      manifest = JSON.parse(manifestString);
    } catch (error) {
      // JSON.parse only throws SyntaxError per doc, which is a subsclass of Error
      throw new JSONSyntaxError(manifestPath, error as Error, actionName);
    }
    AadManifestHelper.processRequiredResourceAccessInManifest(manifest);
    return manifest;
  } finally {
    if (generatedNewPermissionId) {
      // restore environment variable to avoid impact to other code
      delete process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID;
    }
  }
}
