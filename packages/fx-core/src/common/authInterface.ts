// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface AuthBasicParameters {
  apis: string[];
}

export interface OAuthParameters extends AuthBasicParameters {
  authorizationUrl: string;
  tokenUrl: string;
  refreshUrl?: string;
  scopes: { [scope: string]: string };
  enablePKCE?: boolean;
}

export interface ApiKeyParameters extends AuthBasicParameters {
  in: "header" | "query";
  name: string;
}

export type AuthParameters = AuthBasicParameters | OAuthParameters | ApiKeyParameters;
