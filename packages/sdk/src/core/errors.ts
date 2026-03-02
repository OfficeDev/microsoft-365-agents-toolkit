// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Error code to trace the error types.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export enum ErrorCode {
  /**
   * Invalid parameter error.
   */
  InvalidParameter = "InvalidParameter",

  /**
   * Invalid configuration error.
   */
  InvalidConfiguration = "InvalidConfiguration",

  /**
   * Invalid certificate error.
   */
  InvalidCertificate = "InvalidCertificate",

  /**
   * Internal error.
   */
  InternalError = "InternalError",

  /**
   * Channel is not supported error.
   */
  ChannelNotSupported = "ChannelNotSupported",

  /**
   * Failed to retrieve sso token
   */
  FailedToRetrieveSsoToken = "FailedToRetrieveSsoToken",

  /**
   * Failed to process sso handler
   */
  FailedToProcessSsoHandler = "FailedToProcessSsoHandler",

  /**
   * Cannot find command
   */
  CannotFindCommand = "CannotFindCommand",

  /**
   * Failed to run sso step
   */
  FailedToRunSsoStep = "FailedToRunSsoStep",

  /**
   * Failed to run dedup step
   */
  FailedToRunDedupStep = "FailedToRunDedupStep",

  /**
   * Sso activity handler is undefined
   */
  SsoActivityHandlerIsUndefined = "SsoActivityHandlerIsUndefined",

  /**
   * Runtime is not supported error.
   */
  RuntimeNotSupported = "RuntimeNotSupported",

  /**
   * User failed to finish the AAD consent flow failed.
   */
  ConsentFailed = "ConsentFailed",

  /**
   * The user or administrator has not consented to use the application error.
   */
  UiRequiredError = "UiRequiredError",

  /**
   * Token is not within its valid time range error.
   */
  TokenExpiredError = "TokenExpiredError",

  /**
   * Call service (AAD or simple authentication server) failed.
   */
  ServiceError = "ServiceError",

  /**
   * Operation failed.
   */
  FailedOperation = "FailedOperation",

  /**
   * Invalid response error.
   */
  InvalidResponse = "InvalidResponse",

  /**
   * Authentication info already exists error.
   */
  AuthorizationInfoAlreadyExists = "AuthorizationInfoAlreadyExists",
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 * @internal
 */
export class ErrorMessage {
  // InvalidConfiguration Error
  static readonly InvalidConfiguration = "{0} in configuration is invalid: {1}.";
  static readonly ConfigurationNotExists = "Configuration does not exist. {0}";
  static readonly ResourceConfigurationNotExists = "{0} resource configuration does not exist.";
  static readonly MissingResourceConfiguration =
    "Missing resource configuration with type: {0}, name: {1}.";
  static readonly AuthenticationConfigurationNotExists =
    "Authentication configuration does not exist.";

  // RuntimeNotSupported Error
  static readonly BrowserRuntimeNotSupported = "{0} is not supported in browser.";
  static readonly NodejsRuntimeNotSupported = "{0} is not supported in Node.";

  // Internal Error
  static readonly FailToAcquireTokenOnBehalfOfUser =
    "Failed to acquire access token on behalf of user: {0}";

  // ChannelNotSupported Error
  static readonly OnlyMSTeamsChannelSupported = "{0} is only supported in MS Teams Channel";

  static readonly FailedToProcessSsoHandler = "Failed to process sso handler: {0}";

  // FailedToRetrieveSsoToken Error
  static readonly FailedToRetrieveSsoToken =
    "Failed to retrieve sso token, user failed to finish the AAD consent flow.";

  // CannotFindCommand Error
  static readonly CannotFindCommand = "Cannot find command: {0}";

  static readonly FailedToRunSsoStep = "Failed to run dialog to retrieve sso token: {0}";

  static readonly FailedToRunDedupStep = "Failed to run dialog to remove duplicated messages: {0}";

  // SsoActivityHandlerIsUndefined Error
  static readonly SsoActivityHandlerIsNull =
    "Sso command can only be used or added when sso activity handler is not undefined";

  // AuthorizationInfoError
  static readonly AuthorizationHeaderAlreadyExists = "Authorization header already exists!";
  static readonly BasicCredentialAlreadyExists = "Basic credential already exists!";
  // InvalidParameter Error
  static readonly EmptyParameter = "Parameter {0} is empty";
  static readonly DuplicateHttpsOptionProperty =
    "Axios HTTPS agent already defined value for property {0}";
  static readonly DuplicateApiKeyInHeader =
    "The request already defined api key in request header with name {0}.";
  static readonly DuplicateApiKeyInQueryParam =
    "The request already defined api key in query parameter with name {0}.";
  static readonly OnlySupportInQueryActivity =
    "The handleMessageExtensionQueryWithToken only support in handleTeamsMessagingExtensionQuery with composeExtension/query type.";
  static readonly OnlySupportInLinkQueryActivity =
    "The handleMessageExtensionLinkQueryWithSSO only support in handleTeamsAppBasedLinkQuery with composeExtension/queryLink type.";
}

/**
 * Error class with code and message thrown by the SDK.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export class ErrorWithCode extends Error {
  /**
   * Error code
   *
   * @readonly
   */
  code: string | undefined;

  /**
   * Constructor of ErrorWithCode.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   * @param {string} message - error message.
   * @param {ErrorCode} code - error code.
   */
  constructor(message?: string, code?: ErrorCode) {
    if (!code) {
      super(message);
      return this;
    }

    super(message);
    Object.setPrototypeOf(this, ErrorWithCode.prototype);
    this.name = `${new.target.name}.${code}`;
    this.code = code;
  }
}
