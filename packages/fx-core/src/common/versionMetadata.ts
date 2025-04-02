// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import { EOL } from "os";

export const MetadataV3 = {
  projectVersion: "1.0.0",
  unSupprotVersion: "2.0.0",
  platformVersion: {
    vs: "17.5.x.x",
    vsc: "5.x.x",
    cli: "2.x.x",
    cli_help: "2.x.x",
  },
  configFile: "teamsapp.yml",
  localConfigFile: "teamsapp.local.yml",
  testToolConfigFile: "teamsapp.testtool.yml",
  sandboxConfigFile: "teamsapp.sandbox.yml",
  defaultEnvironmentFolder: "env",
  envFilePrefix: ".env",
  secretFileSuffix: "user",
  projectId: "projectId",
  teamsManifestFolder: "appPackage",
  teamsManifestFileName: "manifest.json",
  aadManifestFileName: "aad.manifest.json",
  v3UpgradeWikiLink: "https://aka.ms/teams-toolkit-5.0-upgrade",
  secretFileComment:
    "# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project." +
    EOL,
  secretComment:
    "# Secrets. Keys prefixed with `SECRET_` will be masked in Teams Toolkit logs." + EOL,
  envFileDevComment:
    "# This file includes environment variables that will be committed to git by default." + EOL,
  envFileLocalComment:
    "# This file includes environment variables that can be committed to git. It's gitignored by default because it represents your local development environment." +
    EOL,
};

export const MetadataV2 = {
  projectVersion: "2.0.0",
  projectMaxVersion: "2.1.0",
  platformVersion: {
    vs: "17.4.x.x",
    vsc: "4.x.x",
    cli: "1.x.x",
    cli_help: "1.x.x",
  },
  configFolder: ".fx",
  stateFolder: "states",
  userdataSuffix: "userdata",
  configFile: "projectSettings.json",
  updateToolkitLink: "https://aka.ms/update-teams-toolkit",
  aadTemplateFileName: "aad.template.json",
};

export const MetadataV3Abandoned = {
  projectVersion: "3.0.0",
  configFolder: "teamsfx",
  configFile: "settings.json",
};

export enum VersionState {
  // project version compatible
  compatible = 0,
  // project version outdated, project should upgrade
  upgradeable = -1,
  // project version ahead, need update toolkit
  unsupported = 1,
}

export enum VersionSource {
  projectSettings, // for v2 project
  teamsapp, // for v3 project
  unknown,
  settings, // for abandoned v3 project
}
export interface VersionInfo {
  version: string;
  source: VersionSource;
}
