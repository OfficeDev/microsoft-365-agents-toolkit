// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import * as util from "util";
import * as constants from "../../constants";
import { Browser } from "./constants";

export function WorkspaceNotSupported(workspaceFolder: string): UserError {
  return new UserError(
    constants.cliSource,
    "WorkspaceNotSupported",
    `Workspace '${workspaceFolder}' is not supported.`
  );
}

export function ExclusiveLocalRemoteOptions(): UserError {
  return new UserError(
    constants.cliSource,
    "ExclusiveLocalRemoteOptions",
    "Options --local and --remote are exclusive."
  );
}

export function RequiredPathNotExists(path: string): UserError {
  return new UserError(
    constants.cliSource,
    "RequiredPathNotExists",
    `Required path '${path}' does not exist.`
  );
}

export function TaskFailed(taskTitle: string): UserError {
  let words = taskTitle.split(" ");
  words = words.map((word) => {
    return word.charAt(0).toUpperCase() + word.slice(1);
  });
  return new UserError(
    constants.cliSource,
    `${words.join("")}Failed`,
    `Task '${taskTitle}' failed.`
  );
}

export function PreviewCommandFailed(fxErrors: FxError[]): UserError {
  const errors = fxErrors.map((error) => {
    return `${error.source}.${error.name}`;
  });
  return new UserError(
    constants.cliSource,
    "PreviewCommandFailed",
    `The preview command failed: ${errors.join(", ")}.`
  );
}

export function TeamsAppIdNotExists(): UserError {
  return new UserError(constants.cliSource, "TeamsAppIdNotExists", "App id does not exist.");
}

export function PortsAlreadyInUse(portsInUse: number[]): UserError {
  const message =
    portsInUse.length > 1
      ? util.format(
          "Ports: %s are already in use. Close these ports and try again.",
          portsInUse.join(", ")
        )
      : util.format("Port: %s is already in use. Close this port and try again.", portsInUse[0]);
  return new UserError(constants.cliSource, "PortsAlreadyInUse", message);
}

export function PreviewWithoutProvision(): UserError {
  return new UserError(
    constants.cliSource,
    "PreviewWithoutProvision",
    "Provision and deploy commands are required before preview from remote."
  );
}

export function MissingProgrammingLanguageSetting(): UserError {
  return new UserError(
    constants.cliSource,
    "MissingProgrammingLanguage",
    "The programmingLanguage config is missing in project settings."
  );
}

export function OpeningBrowserFailed(browser: Browser): UserError {
  return new UserError(
    constants.cliSource,
    "OpeningBrowserFailed",
    `Failed to open ${browser} browser. Check if ${browser} exists on your system.`
  );
}

export function OpeningTeamsDesktopClientFailed(): UserError {
  return new UserError(
    constants.cliSource,
    "OpeningTeamsDesktopClientFailed",
    `Failed to open Teams desktop client. Check if Teams exists on your system.`
  );
}

export function NoUrlForSPFxRemotePreview(): UserError {
  return new UserError(
    constants.cliSource,
    "NoUrlForSPFxRemotePreview",
    "SPFx remote preview need your SharePoint site url, pls input sharepoint-site parameter."
  );
}

export function InvalidSharePointSiteURL(error: Error): UserError {
  return new UserError(constants.cliSource, "InvalidSharePointSiteURL", error.message);
}

export function DependencyCheckerFailed(): SystemError {
  return new SystemError(
    constants.cliSource,
    "DependencyCheckerFailed",
    "dependency checker failed."
  );
}

export function PrerequisitesValidationNodejsError(
  error: string | Error,
  helpLink?: string
): UserError {
  return new UserError({
    source: constants.cliSource,
    name: "PrerequisitesValidationNodejsError",
    helpLink,
    message: error instanceof Error ? error.message : (error as string),
  });
}

export function PrerequisitesValidationM365AccountError(
  error: string | Error,
  helpLink?: string
): UserError {
  return new UserError({
    source: constants.cliSource,
    name: "PrerequisitesValidationM365AccountError",
    helpLink,
    message: error instanceof Error ? error.message : (error as string),
  });
}

export function NpmInstallFailed(): UserError {
  return new UserError(constants.cliSource, "NpmInstallFailed", "Npm install failed.");
}

export function CannotDetectRunCommand(): UserError {
  return new UserError(
    constants.cliSource,
    "CannotDetectRunCommand",
    "Cannot detect run command by project type. Parameter '--run-command' is required."
  );
}
