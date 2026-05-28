// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DepsType, MosServiceScope } from "@microsoft/teamsfx-core";

export enum Checker {
  M365Account = "Microsoft 365 Account",
  CopilotAccess = "Copilot Access",
  Ports = "ports occupancy",
}

export const DepsDisplayName = {
  [DepsType.LtsNode]: "Node.js",
  [DepsType.ProjectNode]: "Node.js",
};

export enum ResultStatus {
  success = "success",
  warn = "warn",
  failed = "failed",
}

export const ProgressMessage = Object.freeze({
  [Checker.M365Account]: `Checking ${Checker.M365Account}`,
  [Checker.CopilotAccess]: `Checking ${Checker.CopilotAccess}`,
  [Checker.Ports]: `Checking ${Checker.Ports}`,
  [DepsType.LtsNode]: `Checking ${DepsDisplayName[DepsType.LtsNode]}`,
  [DepsType.ProjectNode]: `Checking ${DepsDisplayName[DepsType.ProjectNode]}`,
});

export const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
