// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs, { CopyOptions, WriteFileOptions } from "fs-extra";
import path from "path";
import yaml from "yaml";

import { Platform } from "@microsoft/teamsfx-api";

import { envUtil } from "../utils/envUtil";
import { pathUtils } from "../utils/pathUtils";

export const backupFolder = ".backup";

export class DeclarativeAgentBotContext {
  private modifiedPaths: string[] = [];
  platform: Platform;
  env = "";
  projectPath = "";
  backupPath = "";
  teamsManifestPath = "";
  agentManifestPath = "";
  multiTenant: boolean;
  isProvisioned = false;
  agentId?: string;
  teamsBotId?: string;

  static async create(
    platform: Platform,
    env: string,
    projectPath: string,
    teamsManifestPath: string,
    agentManifestPath: string,
    multiTenant: boolean
  ): Promise<DeclarativeAgentBotContext> {
    const context = new DeclarativeAgentBotContext(
      platform,
      env,
      projectPath,
      teamsManifestPath,
      agentManifestPath,
      multiTenant
    );
    await fs.ensureDir(context.backupPath);
    return context;
  }

  private constructor(
    platform: Platform,
    env: string,
    projectPath: string,
    teamsManifsetPath: string,
    agentManifestPath: string,
    multiTenant: boolean
  ) {
    this.platform = platform;
    this.env = env;
    this.projectPath = projectPath;
    this.teamsManifestPath = teamsManifsetPath;
    this.agentManifestPath = agentManifestPath;
    this.backupPath = path.join(this.projectPath, backupFolder);
    this.multiTenant = multiTenant;
  }

  async backup(_path: string): Promise<boolean> {
    const srcPath = path.join(this.projectPath, _path);
    if (await fs.pathExists(srcPath)) {
      await fs.copy(srcPath, path.join(this.projectPath, backupFolder, _path));
      return true;
    }
    return false;
  }

  addModifiedPath(path: string): void {
    if (!this.modifiedPaths.includes(path)) {
      this.modifiedPaths.push(path);
    }
  }

  async cleanModifiedPaths(): Promise<void> {
    for (const modifiedPath of this.modifiedPaths.reverse()) {
      await this.fsRemove(modifiedPath);
    }
    this.modifiedPaths.length = 0;
  }

  async fsCopy(src: string, dest: string, options?: CopyOptions): Promise<void> {
    await fs.copy(path.join(this.projectPath, src), path.join(this.projectPath, dest), options);
    this.addModifiedPath(dest);
  }

  async fsRemove(_path: string): Promise<void> {
    return await fs.remove(path.join(this.projectPath, _path));
  }

  async fsWriteFile(
    file: string,
    data: unknown,
    options?: WriteFileOptions | string
  ): Promise<void> {
    await fs.writeFile(path.join(this.projectPath, file), data, options);
    this.addModifiedPath(file);
  }

  async writeEnv(key: string, value: string): Promise<void> {
    const ymlFilePath = pathUtils.getYmlFilePath(this.projectPath, "dev");
    const ymlContent = await fs.readFile(ymlFilePath, "utf-8");
    const yamlObj = yaml.parse(ymlContent);
    const folderPath = yamlObj.environmentFolderPath?.toString() || "./env";
    const envFilePath = path.join(folderPath, `.env.${this.env}`);

    await this.backup(envFilePath);
    await envUtil.writeEnv(this.projectPath, this.env, { [key]: value });
    this.addModifiedPath(envFilePath);
  }

  async cleanBackup(): Promise<void> {
    await this.fsRemove(backupFolder);
  }

  async restoreBackup(): Promise<void> {
    const paths = await fs.readdir(this.backupPath);
    await Promise.all(
      paths.map(async (_path) => {
        await fs.copy(path.join(this.backupPath, _path), path.join(this.projectPath, _path));
      })
    );
  }
}
