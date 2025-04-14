// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import yaml from "yaml";
import { MetadataV3, MetadataV4 } from "../../common/versionMetadata";
import { environmentNameManager } from "../../core/environmentName";
import { MissingRequiredFileError } from "../../error/common";

class PathUtils {
  getYmlFilePath(projectPath: string, env?: string, silent = false): string | undefined {
    if (process.env.TEAMSFX_CONFIG_FILE_PATH) return process.env.TEAMSFX_CONFIG_FILE_PATH;
    const envName = env || process.env.TEAMSFX_ENV || "dev";
    const ymlPathV4 = path.join(
      projectPath,
      envName === environmentNameManager.getLocalEnvName()
        ? MetadataV4.localConfigFile
        : envName === environmentNameManager.getTestToolEnvName()
        ? MetadataV4.testToolConfigFile
        : envName === environmentNameManager.getSandboxEnvName()
        ? MetadataV4.sandboxConfigFile
        : MetadataV4.configFile
    );
    if (fs.pathExistsSync(ymlPathV4)) {
      return ymlPathV4;
    }
    const ymlPathV3 = path.join(
      projectPath,
      envName === environmentNameManager.getLocalEnvName()
        ? MetadataV3.localConfigFile
        : envName === environmentNameManager.getTestToolEnvName()
        ? MetadataV3.testToolConfigFile
        : envName === environmentNameManager.getSandboxEnvName()
        ? MetadataV3.sandboxConfigFile
        : MetadataV3.configFile
    );
    if (silent) return undefined;
    if (environmentNameManager.isRemoteEnvironment(envName)) {
      throw new MissingRequiredFileError("core", "", ymlPathV3);
    } else {
      throw new MissingRequiredFileError("core", "Debug ", ymlPathV3);
    }
  }
  async getEnvFolderPath(projectPath: string): Promise<Result<string | undefined, FxError>> {
    const ymlFilePath = this.getYmlFilePath(projectPath, "dev") as string;
    const ymlContent = await fs.readFile(ymlFilePath, "utf-8");
    const yamlObj = yaml.parse(ymlContent);
    const folderPath = yamlObj.environmentFolderPath?.toString() || "./env";
    const envFolderPath = path.isAbsolute(folderPath)
      ? folderPath
      : path.join(projectPath, folderPath);
    if (!(await fs.pathExists(envFolderPath))) return ok(undefined);
    return ok(envFolderPath);
  }
  async getEnvFilePath(
    projectPath: string,
    env: string
  ): Promise<Result<string | undefined, FxError>> {
    const envFolderPathRes = await this.getEnvFolderPath(projectPath);
    if (envFolderPathRes.isErr()) return err(envFolderPathRes.error);
    const folderPath = envFolderPathRes.value;
    if (!folderPath) return ok(undefined);
    const envFilePath = path.join(folderPath, `.env.${env}`);
    return ok(envFilePath);
  }
  resolveFilePath(projectPath: string, absoluteOrRelativePath?: string): string {
    if (!absoluteOrRelativePath) return projectPath;
    if (path.isAbsolute(absoluteOrRelativePath)) return absoluteOrRelativePath;
    return path.join(projectPath, absoluteOrRelativePath);
  }
}

export const pathUtils = new PathUtils();
