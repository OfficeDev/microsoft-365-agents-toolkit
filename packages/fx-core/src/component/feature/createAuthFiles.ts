// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, Inputs, ok, Result, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { getLocalizedString } from "../../common/localizeUtils";
import { unzip } from "../generator/utils";
import { FileNotFoundError } from "../../error/common";
import { getTemplatesFolder } from "../../folder";
import { AddSsoParameters, SolutionError, SolutionSource } from "../constants";

export async function createAuthFiles(input: Inputs): Promise<Result<unknown, FxError>> {
  const projectPath = input.projectPath;
  if (!projectPath) {
    const e = new SystemError(
      SolutionSource,
      SolutionError.InvalidProjectPath,
      getLocalizedString("core.addSsoFiles.emptyProjectPath")
    );
    return err(e);
  }

  const projectFolderExists = await fs.pathExists(projectPath);
  if (!projectFolderExists) {
    const e = new FileNotFoundError("aad", projectPath);
    return err(e);
  }

  const authFolder = path.join(projectPath, AddSsoParameters.V3AuthFolder);
  try {
    const authFolderExists = await fs.pathExists(authFolder);
    if (!authFolderExists) {
      await fs.ensureDir(authFolder);
    }

    const templateFolder = getTemplatesFolder();
    const v3TemplateFolder = path.join(
      templateFolder,
      AddSsoParameters.filePath,
      AddSsoParameters.V3
    );

    const sampleZip = new AdmZip();
    sampleZip.addLocalFolder(v3TemplateFolder);
    await unzip(sampleZip, authFolder);
    return ok(undefined);
  } catch (error: any) {
    if (await fs.pathExists(authFolder)) {
      await fs.remove(authFolder);
    }
    const e = new SystemError(
      SolutionSource,
      SolutionError.FailedToCreateAuthFiles,
      getLocalizedString("core.addSsoFiles.FailedToCreateAuthFiles", error.message)
    );
    return err(e);
  }
}
