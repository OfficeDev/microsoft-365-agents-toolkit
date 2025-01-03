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

/**
 * Represents the agent hint data structure used for debugging Teams apps
 */
export interface AgentHintData {
  /**
   * The M365 app ID from the environment variables
   */
  id: string;

  /**
   * The scenario identifier for the debug session
   * Currently only supports "launchcopilotextension"
   */
  scenario: "launchcopilotextension";

  /**
   * Additional properties for the agent hint
   */
  properties: {
    /**
     * Timestamp when the debug session was initiated
     */
    clickTimestamp: string;
  };

  /**
   * Version number of the agent hint format
   * Currently set to 1
   */
  version: 1;
}
