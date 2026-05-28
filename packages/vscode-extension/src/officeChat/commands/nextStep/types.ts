// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { MachineStatus, ProjectActionStatus } from "../../../chat/commands/nextstep/types";

export interface OfficeWholeStatus {
  machineStatus: MachineStatus;
  projectOpened?: {
    path: string; // the path of the opened app
    projectId?: string; // the project id of the opened app, it is from teamsapp.yml
    codeModifiedTime: {
      source: Date; // the time when the source code is modified
      infra: Date; // the time when the infra is modified
    };
    actionStatus: ProjectActionStatus;
    readmeContent?: string; // the content of the readme file
    launchJSONContent?: string; // the content of the .vscode/launch.json
    nodeModulesExist?: boolean; // if the node_modules folder exists
    isNodeInstalled?: boolean; // if node.js is installed
  };
}
