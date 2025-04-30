// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export { ErrorWithCode, ErrorCode } from "./core/errors";

export { AppCredential } from "./credential/appCredential";
export { OnBehalfOfUserCredential } from "./credential/onBehalfOfUserCredential";
export { TeamsUserCredential } from "./credential/teamsUserCredential";

export { TeamsBotSsoPrompt, TeamsBotSsoPromptSettings } from "./bot/teamsBotSsoPrompt";
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
export { BasicAuthProvider } from "./apiClient/basicAuthProvider";
export { ApiKeyProvider, ApiKeyLocation } from "./apiClient/apiKeyProvider";
export {
  CertificateAuthProvider,
  createPemCertOption,
  createPfxCertOption,
} from "./apiClient/certificateAuthProvider";

export {
  AdaptiveCardResponse,
  CommandMessage,
  CommandOptions,
  CardActionOptions,
  NotificationTarget,
  NotificationTargetType,
  ConversationReferenceStore,
  ConversationReferenceStoreAddOptions,
  PagedData,
  InvokeResponseErrorCode,
  TriggerPatterns,
  TeamsFxAdaptiveCardActionHandler,
  TeamsFxBotCommandHandler,
  TeamsFxBotSsoCommandHandler,
  BotSsoExecutionActivityHandler,
  BotSsoConfig,
  BotSsoExecutionDialogHandler,
} from "./conversation/interface";
export { BotSsoExecutionDialog } from "./conversation/sso/botSsoExecutionDialog";
export { MessageBuilder } from "./conversation/messageBuilder";
export { InvokeResponseFactory } from "./conversation/invokeResponseFactory";
export {
  handleMessageExtensionQueryWithSSO,
  handleMessageExtensionLinkQueryWithSSO,
} from "./messageExtension/executeWithSSO";
export { MessageExtensionTokenResponse } from "./messageExtension/teamsMsgExtTokenResponse";

import * as AgentBuilderCloudAdapter from "./conversationWithCloudAdapter/conversationWithCloudAdapter";
export { AgentBuilderCloudAdapter };
