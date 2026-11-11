// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import yaml from "yaml";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";
import { MetadataV3, MetadataV4 } from "../../common/versionMetadata";
import { environmentNameManager } from "../../core/environmentName";
import { MissingRequiredFileError } from "../../error/common";

class PathUtils {
  getAvailableYmlFilePath(
    projectPath: string,
    validatePath?: (filePath: string) => void
  ): string | undefined {
    const possibleEnvs = ["", ".playground", ".local"];
    for (const env of possibleEnvs) {
      const ymlPath = path.join(projectPath, `m365agents${env}.yml`);
      validatePath?.(ymlPath);
      if (fs.pathExistsSync(ymlPath)) return ymlPath;
    }
    return undefined;
  }

  getYmlFilePath(
    projectPath: string,
    env?: string,
    silent = false,
    validatePath?: (filePath: string) => void
  ): string | undefined {
    if (process.env.TEAMSFX_CONFIG_FILE_PATH) {
      validatePath?.(process.env.TEAMSFX_CONFIG_FILE_PATH);
      return process.env.TEAMSFX_CONFIG_FILE_PATH;
    }
    const envName = env || process.env.TEAMSFX_ENV || "dev";
    const ymlPathV4 = path.join(
      projectPath,
      envName === environmentNameManager.getLocalEnvName()
        ? MetadataV4.localConfigFile
        : envName === environmentNameManager.getPlaygroundEnvName()
          ? MetadataV4.testToolConfigFile
          : envName === environmentNameManager.getSandboxEnvName()
            ? MetadataV4.sandboxConfigFile
            : MetadataV4.configFile
    );
    validatePath?.(ymlPathV4);
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
    validatePath?.(ymlPathV3);
    if (fs.pathExistsSync(ymlPathV3)) {
      return ymlPathV3;
    }
    if (featureFlagManager.getBooleanValue(FeatureFlags.GenerateConfigFiles)) {
      const availableYmlFilePath = this.getAvailableYmlFilePath(projectPath, validatePath) || "";
      if (fs.pathExistsSync(availableYmlFilePath)) {
        return availableYmlFilePath;
      }
    }
    if (silent) return undefined;
    if (environmentNameManager.isRemoteEnvironment(envName)) {
      throw new MissingRequiredFileError("core", "", ymlPathV4);
    } else {
      throw new MissingRequiredFileError("core", "Debug ", ymlPathV4);
    }
  }
  async getEnvFolderPath(
    projectPath: string,
    env = "dev",
    validatePath?: (filePath: string) => void
  ): Promise<Result<string | undefined, FxError>> {
    const ymlFilePath = this.getYmlFilePath(projectPath, env, false, validatePath);
    if (ymlFilePath === undefined) {
      return err(new MissingRequiredFileError("core", "", "m365agents.yml"));
    }
    const ymlContent = await fs.readFile(ymlFilePath, "utf-8");
    const yamlObj = yaml.parse(ymlContent);
    const folderPath = yamlObj.environmentFolderPath?.toString() || "./env";
    const envFolderPath = path.isAbsolute(folderPath)
      ? folderPath
      : path.join(projectPath, folderPath);
    validatePath?.(envFolderPath);
    if (!(await fs.pathExists(envFolderPath))) return ok(undefined);
    return ok(envFolderPath);
  }
  async getEnvFilePath(
    projectPath: string,
    env: string,
    validatePath?: (filePath: string) => void
  ): Promise<Result<string | undefined, FxError>> {
    const envFolderPathRes = await this.getEnvFolderPath(projectPath, env, validatePath);
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
