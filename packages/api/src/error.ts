// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface FxError extends Error {
  /**
   * Custom error details.
   */
  innerError?: any;
  data?: Record<string, unknown>;
  /**
   * Source name of error. (plugin name, eg: tab-scaffhold-plugin)
   */
  source: string;
  /**
   * Time of error.
   */
  timestamp: Date;

  userData?: any;

  categories?: string[];

  telemetryProperties?: Record<string, string>;

  /**
   * recommended operation for user to fix the error
   * e.g. "debug-in-test-tool"
   */
  recommendedOperation?: string;
  /**
   * whether to skip process (such as mask secret tokens) in telemetry collection
   */
  skipProcessInTelemetry?: boolean;
}
export interface ErrorOptionBase {
  source?: string;
  name?: string;
  message?: string;
  error?: Error;
  userData?: any;
  data?: Record<string, unknown>;
  displayMessage?: string;
  categories?: string[];
  telemetryProperties?: Record<string, string>;
  /**
   * whether to skip process (such as mask secret tokens) in telemetry collection
   */
  skipProcessInTelemetry?: boolean;
}

export interface UserErrorOptions extends ErrorOptionBase {
  helpLink?: string;
}

export interface SystemErrorOptions extends ErrorOptionBase {
  issueLink?: string;
}
/**
 * Users can recover by themselves, e.g., users input invalid app names.
 */
export class UserError extends Error implements FxError {
  /**
   * Custom error details .
   */
  innerError?: any;
  /**
   * Source name of error. (plugin name, eg: tab-scaffold-plugin)
   */
  source: string;
  /**
   * Time of error.
   */
  timestamp: Date;
  /**
   * A wiki website that shows mapping relationship between error names, descriptions, and fix solutions.
   */
  helpLink?: string;
  /**
   * data that only be reported to github issue  manually by user and will not be reported as telemetry data
   */
  userData?: string;
  data?: Record<string, unknown>;
  /**
   * message show in the UI
   */
  displayMessage?: string;

  categories?: string[];

  telemetryProperties?: Record<string, string>;

  /**
   * whether to skip process (such as mask secret tokens) in telemetry collection
   */
  skipProcessInTelemetry?: boolean;

  /**
   * recommended operation for user to fix the error
   * e.g. "debug-in-test-tool"
   */
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

    // message
    const message = option.message || option.error?.message;
    super(message);

    //name
    this.name = option.name || new.target.name;

    //source
    this.source = option.source || "unknown";

    //stack
    Error.captureStackTrace(this, new.target);

    //prototype
    Object.setPrototypeOf(this, new.target.prototype);

    //innerError
    this.innerError = option.error;
    this.data = option.data;

    //other fields
    this.helpLink = option.helpLink;
    this.userData = option.userData;
    this.displayMessage = option.displayMessage;
    this.timestamp = new Date();
    this.categories = option.categories;
    this.skipProcessInTelemetry = option.skipProcessInTelemetry;
    this.telemetryProperties = option.telemetryProperties;
  }
}

/**
 * Users cannot handle it by themselves.
 */
export class SystemError extends Error implements FxError {
  /**
   * Custom error details.
   */
  innerError?: any;
  data?: Record<string, unknown>;
  /**
   * Source name of error. (plugin name, eg: tab-scaffold-plugin)
   */
  source: string;
  /**
   * Time of error.
   */
  timestamp: Date;
  /**
   * A github issue page where users can submit a new issue.
   */
  issueLink?: string;

  /**
   * data that only be reported to github issue  manually by user and will not be reported as telemetry data
   */
  userData?: string;
  /**
   * message show in the UI
   */
  displayMessage?: string;

  categories?: string[];

  telemetryProperties?: Record<string, string>;

  /**
   * whether to skip process (such as mask secret tokens) in telemetry collection
   */
  skipProcessInTelemetry?: boolean;

  /**
   * recommended operation for user to fix the error
   * e.g. "debug-in-test-tool"
   */
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

    // message
    const message = option.message || option.error?.message;
    super(message);

    //name
    this.name = option.name || new.target.name;

    //source
    this.source = option.source || "unknown";

    //stack
    Error.captureStackTrace(this, new.target);

    //prototype
    Object.setPrototypeOf(this, new.target.prototype);

    //innerError
    this.innerError = option.error;
    this.data = option.data;

    //other fields
    this.issueLink = option.issueLink;
    this.userData = option.userData;
    this.displayMessage = option.displayMessage;
    this.timestamp = new Date();
    this.categories = option.categories;
    this.skipProcessInTelemetry = option.skipProcessInTelemetry;
    this.telemetryProperties = option.telemetryProperties;
  }
}
