// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError, UserError, UserErrorOptions } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { ErrorCategory } from "./types";

export const NodejsNotFoundHelpLink = "https://aka.ms/teamsfx-node";
export const NodejsNotRecommendedHelpLink = "https://aka.ms/teamsfx-node";
export const NodejsNotLtsHelpLink = "https://aka.ms/teamsfx-node";

export class PortsConflictError extends UserError {
  constructor(ports: number[], occupiedPorts: number[], source?: string) {
    const key = "error.dep.PortsConflictError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "PortsConflictError",
      message: getDefaultString(key, ports.join(", "), occupiedPorts.join(", ")),
      displayMessage: getLocalizedString(key, ports.join(", "), occupiedPorts.join(", ")),
      categories: [ErrorCategory.Internal],
      telemetryProperties: {
        ports: ports.join(", "),
        "occupied-ports": occupiedPorts.join(", "),
      },
    };
    super(errorOptions);
  }
}

export class SideloadingDisabledError extends UserError {
  constructor(source?: string) {
    const key = "error.dep.SideloadingDisabledError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "SideloadingDisabledError",
      message: getDefaultString(key),
      displayMessage: getLocalizedString(key),
      categories: [ErrorCategory.External],
    };
    super(errorOptions);
  }
}

export class CopilotDisabledError extends UserError {
  constructor(source?: string) {
    const key = "error.dep.CopilotDisabledError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "CopilotDisabledError",
      message: getDefaultString(key),
      displayMessage: getLocalizedString(key),
      categories: [ErrorCategory.External],
    };
    super(errorOptions);
  }
}

export class NodejsNotLtsError extends UserError {
  constructor(version: string, source?: string) {
    const key = "error.dep.NodejsNotLtsError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "NodejsNotLtsError",
      message: getDefaultString(key, version),
      displayMessage: getLocalizedString(key, version),
      categories: [ErrorCategory.Internal],
      helpLink: NodejsNotLtsHelpLink,
      telemetryProperties: {
        "nodejs-version": version,
      },
    };
    super(errorOptions);
  }
}

export class NodejsNotFoundError extends UserError {
  constructor(source?: string) {
    const key = "error.dep.NodejsNotFoundError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "NodejsNotFoundError",
      message: getDefaultString(key),
      displayMessage: getLocalizedString(key),
      categories: [ErrorCategory.Internal],
      helpLink: NodejsNotFoundHelpLink,
    };
    super(errorOptions);
  }
}

export class NodejsNotRecommendedError extends UserError {
  constructor(version: string, recommendVersion: string, source?: string) {
    const key = "error.dep.NodejsNotRecommendedError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "NodejsNotRecommendedError",
      message: getDefaultString(key, version, recommendVersion),
      displayMessage: getLocalizedString(key, version, recommendVersion),
      categories: [ErrorCategory.Internal],
      telemetryProperties: {
        "nodejs-version": version,
        "nodejs-version-recommended": recommendVersion,
      },
      helpLink: NodejsNotRecommendedHelpLink,
    };
    super(errorOptions);
  }
}

export class VxTestAppInvalidInstallOptionsError extends UserError {
  constructor(source?: string) {
    const key = "error.dep.VxTestAppInvalidInstallOptionsError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "VxTestAppInvalidInstallOptionsError",
      message: getDefaultString(key),
      displayMessage: getLocalizedString(key),
      categories: [ErrorCategory.Internal],
    };
    super(errorOptions);
  }
}

export class VxTestAppValidationError extends UserError {
  constructor(source?: string) {
    const key = "error.dep.VxTestAppValidationError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "VxTestAppValidationError",
      message: getDefaultString(key),
      displayMessage: getLocalizedString(key),
      categories: [ErrorCategory.Internal],
    };
    super(errorOptions);
  }
}

export class DepsCheckerError extends UserError {
  constructor(message: string, helpLink: string, source?: string) {
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "DepsCheckerError",
      message: message,
      displayMessage: message,
      categories: [ErrorCategory.External],
      helpLink: NodejsNotRecommendedHelpLink,
    };
    super(errorOptions);
  }
}

export class FindProcessError extends SystemError {
  constructor(error: Error, source?: string) {
    const key = "error.dep.FindProcessError";
    const errorOptions: UserErrorOptions = {
      source: source || "core",
      name: "FindProcessError",
      error: error,
      message: getDefaultString(key, error.message),
      displayMessage: getLocalizedString(key, error.message),
      categories: [ErrorCategory.Internal],
    };
    super(errorOptions);
  }
}
