// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface FxError extends Error {
  innerError?: any;
  source: string;
  timestamp: Date;
  userData?: any;
  categories?: string[];
  telemetryProperties?: Record<string, string>;
  recommendedOperation?: string;
  skipProcessInTelemetry?: boolean;
}

export interface ErrorOptionBase {
  source?: string;
  name?: string;
  message?: string;
  error?: Error;
  userData?: any;
  displayMessage?: string;
  categories?: string[];
  telemetryProperties?: Record<string, string>;
  skipProcessInTelemetry?: boolean;
}

export interface UserErrorOptions extends ErrorOptionBase {
  helpLink?: string;
}

export interface SystemErrorOptions extends ErrorOptionBase {
  issueLink?: string;
}

export class UserError extends Error implements FxError {
  innerError?: any;
  source: string;
  timestamp: Date;
  helpLink?: string;
  userData?: string;
  displayMessage?: string;
  categories?: string[];
  telemetryProperties?: Record<string, string>;
  skipProcessInTelemetry?: boolean;
  recommendedOperation?: string;

  constructor(opt: UserErrorOptions);
  constructor(source: string, name: string, message: string, displayMessage?: string);
  constructor(
    param1: string | UserErrorOptions,
    param2?: string,
    param3?: string,
    param4?: string
  ) {
    let option: UserErrorOptions;
    if (typeof param1 === "string") {
      option = {
        source: param1,
        name: param2,
        message: param3,
        displayMessage: param4,
      };
    } else {
      option = param1;
    }
    const message = option.message || option.error?.message;
    super(message);
    this.name = option.name || new.target.name;
    this.source = option.source || "unknown";
    Error.captureStackTrace(this, new.target);
    Object.setPrototypeOf(this, new.target.prototype);
    this.innerError = option.error;
    this.helpLink = option.helpLink;
    this.userData = option.userData;
    this.displayMessage = option.displayMessage;
    this.timestamp = new Date();
    this.categories = option.categories;
    this.skipProcessInTelemetry = option.skipProcessInTelemetry;
    this.telemetryProperties = option.telemetryProperties;
  }
}

export class SystemError extends Error implements FxError {
  innerError?: any;
  source: string;
  timestamp: Date;
  issueLink?: string;
  userData?: string;
  displayMessage?: string;
  categories?: string[];
  telemetryProperties?: Record<string, string>;
  skipProcessInTelemetry?: boolean;
  recommendedOperation?: string;

  constructor(opt: SystemErrorOptions);
  constructor(source: string, name: string, message: string, displayMessage?: string);
  constructor(
    param1: string | SystemErrorOptions,
    param2?: string,
    param3?: string,
    param4?: string
  ) {
    let option: SystemErrorOptions;
    if (typeof param1 === "string") {
      option = {
        source: param1,
        name: param2,
        message: param3,
        displayMessage: param4,
      };
    } else {
      option = param1;
    }
    const message = option.message || option.error?.message;
    super(message);
    this.name = option.name || new.target.name;
    this.source = option.source || "unknown";
    Error.captureStackTrace(this, new.target);
    Object.setPrototypeOf(this, new.target.prototype);
    this.innerError = option.error;
    this.issueLink = option.issueLink;
    this.userData = option.userData;
    this.displayMessage = option.displayMessage;
    this.timestamp = new Date();
    this.categories = option.categories;
    this.skipProcessInTelemetry = option.skipProcessInTelemetry;
    this.telemetryProperties = option.telemetryProperties;
  }
}
