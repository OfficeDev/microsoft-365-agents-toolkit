// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export { ErrorWithCode, ErrorCode } from "./core/errors";

export { AppCredential } from "./credential/appCredential.browser";
export { OnBehalfOfUserCredential } from "./credential/onBehalfOfUserCredential.browser";
export { TeamsUserCredential } from "./credential/teamsUserCredential.browser";

export { TeamsBotSsoPrompt, TeamsBotSsoPromptSettings } from "./bot/teamsBotSsoPrompt.browser";
export { TeamsBotSsoPromptTokenResponse } from "./bot/teamsBotSsoPromptTokenResponse";

export { UserInfo } from "./models/userinfo";
export {
  AppCredentialAuthConfig,
  OnBehalfOfCredentialAuthConfig,
  TeamsUserCredentialAuthConfig,
} from "./models/configuration";
export { GetTeamsUserTokenOptions } from "./models/teamsUserTokenOptions";

export {
  Logger,
  LogLevel,
  LogFunction,
  setLogLevel,
  getLogLevel,
  setLogger,
  setLogFunction,
} from "./util/logger";

export { createApiClient } from "./apiClient/apiClient";
export { AxiosInstance } from "axios";
export { AuthProvider } from "./apiClient/authProvider";
export { BearerTokenAuthProvider } from "./apiClient/bearerTokenAuthProvider";
export { BasicAuthProvider } from "./apiClient/basicAuthProvider.browser";
export { ApiKeyProvider, ApiKeyLocation } from "./apiClient/apiKeyProvider.browser";
export {
  CertificateAuthProvider,
  createPemCertOption,
  createPfxCertOption,
} from "./apiClient/certificateAuthProvider.browser";

export {
  AdaptiveCardResponse,
  CommandMessage,
  CommandOptions,
  CardActionOptions,
  NotificationTarget,
  NotificationTargetType,
  InvokeResponseErrorCode,
  TriggerPatterns,
  TeamsFxAdaptiveCardActionHandler,
  TeamsFxBotCommandHandler,
  TeamsFxBotSsoCommandHandler,
  BotSsoExecutionActivityHandler,
  BotSsoConfig,
  BotSsoExecutionDialogHandler,
} from "./conversation/interface";
export { BotSsoExecutionDialog } from "./conversation/sso/botSsoExecutionDialog.browser";
export { handleMessageExtensionQueryWithSSO } from "./messageExtension/executeWithSSO.browser";
export { MessageExtensionTokenResponse } from "./messageExtension/teamsMsgExtTokenResponse";

import * as AgentBuilderCloudAdapter from "./conversationWithCloudAdapter/conversationWithCloudAdapter.browser";
export { AgentBuilderCloudAdapter };
