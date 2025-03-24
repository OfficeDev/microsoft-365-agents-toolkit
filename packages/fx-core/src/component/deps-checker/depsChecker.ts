// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { UserError } from "@microsoft/teamsfx-api";

export interface DepsChecker {
  getInstallationInfo(installOptions?: InstallOptions): Promise<DependencyStatus>;

  resolve(installOptions?: InstallOptions): Promise<DependencyStatus>;
}

export type DependencyStatus = {
  name: string;
  type: DepsType;
  isInstalled: boolean;
  command: string;
  details: {
    isLinuxSupported: boolean;
    supportedVersions: string[];
    installVersion?: string;
    binFolders?: string[];
  };
  telemetryProperties?: { [key: string]: string };
  error?: UserError;
};

export interface DepsInfo {
  name: string;
  isLinuxSupported: boolean;
  installVersion?: string;
  supportedVersions: string[];
  binFolders?: string[];
  details: Map<string, string>;
}

export enum DepsType {
  LtsNode = "lts-node",
  ProjectNode = "project-node",
  Dotnet = "dotnet",
  FuncCoreTools = "func-core-tools",
  TestTool = "test-tool",
  VxTestApp = "vx-test-app",
}

export interface BaseInstallOptions {
  projectPath?: string;
  version?: string;
}

export interface FuncInstallOptions {
  symlinkDir?: string;
  projectPath: string;
  version: string;
}

export interface TestToolInstallOptions {
  npmExecutablePath?: string;
  releaseType: TestToolReleaseType;
  symlinkDir?: string;
  projectPath: string;
  versionRange: string;
}

export enum TestToolReleaseType {
  Npm = "npm",
  Binary = "binary",
}

export type InstallOptions = BaseInstallOptions | FuncInstallOptions | TestToolInstallOptions;
