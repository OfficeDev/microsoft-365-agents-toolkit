// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export { GraphApiClient } from "./client";

// Re-export types, but exclude PublishedAppDefinition to avoid conflict
// with teamsDevPortal/types (both define it). Consumers that need the
// Graph-specific one import directly from "./graphApi/types".
export {
  AADApplication,
  OAuth2PermissionScope,
  PreAuthorizedApplication,
  OptionalClaim,
  RequiredResourceAccess,
  ResourceAccess,
  PasswordCredential,
  AadOwner,
  SignInAudience,
  BotRegistration,
  BotChannelType,
  PublishingState,
  GRAPH_BASE_URL,
  GRAPH_BETA_URL,
  graphScopes,
  graphAppCatalogScopes,
} from "./types";
