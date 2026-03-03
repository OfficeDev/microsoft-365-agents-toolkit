// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs, { CopyOptions, EnsureOptions, PathLike, WriteFileOptions } from "fs-extra";
import path from "path";
import { MetadataV2 } from "../../../common/versionMetadata";
import { CoreHookContext } from "../../types";
import { Inputs } from "@microsoft/teamsfx-api";

export const backupFolder = ".backup";
// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export interface MigrationContext extends CoreHookContext {
  backup(path: string): Promise<boolean>;
  fsEnsureDir(path: string, options?: EnsureOptions | number): Promise<void>;
  fsCopy(src: string, dest: string, options?: CopyOptions): Promise<void>;
  fsCreateFile(file: string): Promise<void>;
  fsWriteFile(
    file: PathLike | number,
    data: any,
    options?: WriteFileOptions | string
  ): Promise<void>;
  addReport(report: string): void;
  addTelemetryProperties(properties: Record<string, string>): void;
  currentStep?: string;
}

// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export class MigrationContext {
  private modifiedPaths: string[] = [];
  private reports: string[] = [];
  telemetryProperties: Record<string, string> = {};
  backupPath = "";
  projectPath = "";
  isBotValidDomain = false;

  static async create(ctx: CoreHookContext): Promise<MigrationContext> {
    const context = new MigrationContext(ctx);
    await fs.ensureDir(context.backupPath);
    return context;
  }

  private constructor(ctx: CoreHookContext) {
    Object.assign(this, ctx, {});
    this.projectPath = getParameterFromCxt(ctx, "projectPath");
    this.backupPath = path.join(this.projectPath, backupFolder);
  }

  async backup(_path: string): Promise<boolean> {
    const srcPath = path.join(this.projectPath, _path);
    if (await fs.pathExists(srcPath)) {
      await fs.copy(srcPath, path.join(this.backupPath, _path));
      return true;
    }
    return false;
  }

  async fsEnsureDir(_path: string, options?: EnsureOptions | number): Promise<void> {
    const srcPath = path.join(this.projectPath, _path);
    const parentPath = path.dirname(srcPath);
    if (!(await fs.pathExists(parentPath))) {
      await this.fsEnsureDir(path.relative(this.projectPath, parentPath), options);
    }
    if (!(await fs.pathExists(srcPath))) {
      await fs.ensureDir(srcPath, options);
      this.addModifiedPath(_path);
    }
  }

  async fsCopy(src: string, dest: string, options?: CopyOptions): Promise<void> {
    await fs.copy(path.join(this.projectPath, src), path.join(this.projectPath, dest), options);
    this.addModifiedPath(dest);
  }

  async fsCreateFile(file: string): Promise<void> {
    await fs.createFile(path.join(this.projectPath, file));
    this.addModifiedPath(file);
  }

  async fsWriteFile(file: string, data: any, options?: WriteFileOptions | string): Promise<void> {
    await fs.writeFile(path.join(this.projectPath, file), data, options);
    this.addModifiedPath(file);
  }

  addModifiedPath(path: string): void {
    if (!this.modifiedPaths.includes(path)) {
      this.modifiedPaths.push(path);
    }
  }

  getModifiedPaths(): string[] {
    return this.modifiedPaths.slice();
  }

  async cleanModifiedPaths(): Promise<void> {
    for (const modifiedPath of this.modifiedPaths.reverse()) {
      await this.fsRemove(modifiedPath);
    }
    this.modifiedPaths.length = 0;
  }

  async restoreBackup(): Promise<void> {
    const paths = await fs.readdir(this.backupPath);
    await Promise.all(
      paths.map(async (_path) => {
        await fs.copy(path.join(this.backupPath, _path), path.join(this.projectPath, _path));
      })
    );
  }

  async cleanBackup(): Promise<void> {
    await this.fsRemove(backupFolder);
  }

  async removeFxV2(): Promise<void> {
    await this.fsRemove(MetadataV2.configFolder);
  }

  async fsPathExists(_path: string): Promise<boolean> {
    return await fs.pathExists(path.join(this.projectPath, _path));
  }

  async fsRemove(_path: string): Promise<void> {
    return await fs.remove(path.join(this.projectPath, _path));
  }

  addReport(report: string): void {
    this.reports.push(report);
  }

  addTelemetryProperties(properties: Record<string, string>): void {
    this.telemetryProperties = { ...properties, ...this.telemetryProperties };
  }
}

function getParameterFromCxt(ctx: CoreHookContext, key: string, defaultValue?: string): string {
  const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
  const value = (inputs[key] as string) || defaultValue || "";
  return value;
}
