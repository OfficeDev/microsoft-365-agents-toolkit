// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SyncManifestInputs } from "../../../../question";
import { DevPortalAppResponse } from "../../../utils/appInfo";

export interface SyncManifestArgs {
  /**
   * Teams app project path
   */
  projectPath: string;
  /**
   * Environment
   */
  env: string;
  /**
   * Teams app id
   */
  teamsAppId?: string;
}

export interface SyncManifestInputsForVS extends SyncManifestInputs {
  teamsAppFromTdp: DevPortalAppResponse;
}
