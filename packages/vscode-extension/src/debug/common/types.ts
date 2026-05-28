// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { DepsType } from "@microsoft/teamsfx-core";
import { ResultStatus, Checker } from "../depsChecker/prerequisitesCheckerConstants";

export type CheckResult = {
  checker: string;
  result: ResultStatus;
  error?: FxError;
  successMsg?: string;
  warnMsg?: string;
  failureMsg?: string;
};

export type PortCheckerInfo = { checker: Checker.Ports; ports: number[] };

export type PrerequisiteCheckerInfo = {
  checker:
    | Checker
    | Checker.M365Account
    | Checker.CopilotAccess
    | Checker.Ports
    | DepsType.LtsNode
    | DepsType.ProjectNode;
  [key: string]: any;
};

export type PrerequisiteOrderedChecker = {
  info: PrerequisiteCheckerInfo;
  fastFail: boolean;
};
