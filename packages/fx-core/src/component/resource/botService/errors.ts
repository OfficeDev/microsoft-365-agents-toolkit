// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Qianhao Dong <qidon@microsoft.com>
 */
import { ErrorNames } from "./constants";
import { Messages } from "./messages";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

const ErrorType = {
  USER: "User",
  SYSTEM: "System",
} as const;

type ErrorType = (typeof ErrorType)[keyof typeof ErrorType];

type InnerError = HttpError | Error | ErrorWithMessage | ErrorWithCode | unknown;

type HttpError = {
  response: {
    status?: number;
    data?: {
      errorMessage?: string;
      error?: {
        code?: string;
        message?: string;
      };
      errors?: any;
    };
  };
};

type ErrorWithMessage = {
  message: string;
};

type ErrorWithCode = {
  code: string;
};

class PluginError extends Error {
  public name: string;
  public details: [string, string];
  public suggestions: string[];
  public errorType: ErrorType;
  public innerError?: InnerError;
  public helpLink?: string;

  constructor(
    type: ErrorType,
    name: string,
    details: [string, string],
    suggestions: string[],
    innerError?: InnerError,
    helpLink?: string
  ) {
    super(details[0]);
    this.name = name;
    this.details = details;
    this.suggestions = suggestions;
    this.errorType = type;
    this.innerError = innerError;
    this.helpLink = helpLink;
    Object.setPrototypeOf(this, PluginError.prototype);
  }

  genMessage(): string {
    let msg = `${this.details[0]} `;
    if (this.suggestions.length > 0) {
      msg += getDefaultString("plugins.bot.ErrorSuggestions", this.suggestions.join(" "));
    }
    return msg;
  }
  genDisplayMessage(): string {
    let msg = `${this.details[1]} `;
    if (this.suggestions.length > 0) {
      msg += getLocalizedString("plugins.bot.ErrorSuggestions", this.suggestions.join(" "));
    }
    return msg;
  }
}

export class ProvisionError extends PluginError {
  constructor(resource: string, innerError?: InnerError) {
    super(
      ErrorType.USER,
      ErrorNames.PROVISION_ERROR,
      Messages.FailToProvisionSomeResource(resource),
      [Messages.CheckOutputLogAndTryToFix, Messages.RetryTheCurrentStep],
      innerError
    );
  }
}

export class ConfigUpdatingError extends PluginError {
  constructor(configName: string, innerError?: InnerError) {
    super(
      ErrorType.USER,
      ErrorNames.CONFIG_UPDATING_ERROR,
      Messages.FailToUpdateConfigs(configName),
      [Messages.CheckOutputLogAndTryToFix, Messages.RetryTheCurrentStep],
      innerError
    );
  }
}

export class BotRegistrationNotFoundError extends PluginError {
  constructor(botId: string, innerError?: InnerError) {
    super(
      ErrorType.USER,
      ErrorNames.BOT_REGISTRATION_NOTFOUND_ERROR,
      Messages.BotRegistrationNotFoundWith(botId),
      [Messages.CheckOutputLogAndTryToFix],
      innerError
    );
  }
}

export class BotFrameworkNotAllowedToAcquireTokenError extends PluginError {
  constructor() {
    super(
      ErrorType.USER,
      ErrorNames.ACQUIRE_BOT_FRAMEWORK_TOKEN_ERROR,
      Messages.NotAllowedToAcquireBotFrameworkToken(),
      [Messages.CheckOutputLogAndTryToFix]
    );
  }
}

export class BotFrameworkForbiddenResultError extends PluginError {
  constructor(error: any) {
    super(
      ErrorType.USER,
      ErrorNames.FORBIDDEN_RESULT_BOT_FRAMEWORK_ERROR,
      Messages.BotProvisionReturnsForbiddenResult(error),
      [Messages.CheckOutputLogAndTryToFix, Messages.RetryTheCurrentStep]
    );
  }
}

export class BotFrameworkConflictResultError extends PluginError {
  constructor() {
    super(
      ErrorType.USER,
      ErrorNames.CONFLICT_RESULT_BOT_FRAMEWORK_ERROR,
      Messages.BotProvisionReturnsConflictResult(),
      [Messages.CheckOutputLogAndTryToFix, Messages.RetryTheCurrentStep]
    );
  }
}

class PreconditionError extends PluginError {
  constructor(name: string) {
    super(ErrorType.USER, ErrorNames.PRECONDITION_ERROR, Messages.SomethingIsMissing(name), [
      Messages.RetryTheCurrentStep,
    ]);
  }
}

export function CheckThrowSomethingMissing<T>(name: string, value: T | undefined): T {
  if (!value) {
    throw new PreconditionError(name);
  }
  return value;
}
