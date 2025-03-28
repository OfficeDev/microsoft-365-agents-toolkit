// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import { getAbsolutePath } from "../../../utils/common";
import { DriverContext } from "../../interface/commonArgs";
import { CreateOauthArgs } from "../interface/createOauthArgs";
import { FeatureFlags, featureFlagManager } from "../../../../common/featureFlags";
import { OpenAPIV3 } from "openapi-types";
import { isEqual } from "lodash";
import { maxDomainPerOauth, maxSecretLength, minSecretLength } from "./constants";
import { OauthDomainInvalidError } from "../error/oauthDomainInvalid";
import { OauthFailedToGetDomainError } from "../error/oauthFailedToGetDomain";
import { OauthAuthInfoInvalid } from "../error/oauthAuthInfoInvalid";
import { UpdateOauthArgs } from "../interface/updateOauthArgs";
import { OauthAuthMissingInSpec } from "../error/oauthAuthMissingInSpec";

export interface OauthInfo {
  domain?: string[];
  authorizationEndpoint?: string;
  tokenExchangeEndpoint?: string;
  tokenRefreshEndpoint?: string;
  scopes?: string[];
  clientId?: string;
}

interface AuthInfo {
  authorizationUrl: string;
  tokenUrl: string;
  refreshUrl?: string;
  scopes?: string[];
}

export async function getAuthInfo(
  args: CreateOauthArgs | UpdateOauthArgs,
  context: DriverContext,
  actionName: string
): Promise<OauthInfo> {
  if (args.baseUrl) {
    if (args.identityProvider === "MicrosoftEntra") {
      return {
        domain: [args.baseUrl],
      };
    } else if (args.authorizationUrl && args.tokenUrl) {
      return {
        domain: [args.baseUrl],
        authorizationEndpoint: args.authorizationUrl,
        tokenExchangeEndpoint: args.tokenUrl,
        tokenRefreshEndpoint: args.refreshUrl, // optional
        scopes: parseScopes(args.scope), // optional
      };
    }
  }

  let authInfo: OauthInfo = {};
  // when update, baseUrl and apiSpecPath are not required
  if (args.apiSpecPath) {
    authInfo = await getandValidateOauthInfoFromSpec(args, context, actionName);
  }

  if (args.baseUrl) authInfo.domain = [args.baseUrl];
  if (args.authorizationUrl) authInfo.authorizationEndpoint = args.authorizationUrl;
  if (args.tokenUrl) authInfo.tokenExchangeEndpoint = args.tokenUrl;
  if (args.refreshUrl) authInfo.tokenRefreshEndpoint = args.refreshUrl;
  if (args.scope) authInfo.scopes = parseScopes(args.scope);

  return authInfo;
}

async function getandValidateOauthInfoFromSpec(
  args: CreateOauthArgs | UpdateOauthArgs,
  context: DriverContext,
  actionName: string
): Promise<OauthInfo> {
  const absolutePath = getAbsolutePath(args.apiSpecPath!, context.projectPath);
  const parser = new SpecParser(absolutePath, {
    allowAPIKeyAuth: false,
    allowBearerTokenAuth: true,
    allowMultipleParameters: true,
    allowOauth2: true,
    projectType: ProjectType.Copilot,
    allowMissingId: true,
    allowSwagger: true,
    allowMethods: ["get", "post", "put", "delete", "patch", "head", "connect", "options", "trace"],
  });
  const listResult = await parser.list();
  const operations = listResult.APIs.filter((value) => {
    const auth = value.auth;
    return auth && auth.authScheme.type === "oauth2" && auth.name === args.name;
  });

  if (operations.length === 0) {
    throw new OauthAuthMissingInSpec(actionName, args.name);
  }

  let domains: string[] = [];
  if (args.baseUrl) {
    domains = [args.baseUrl];
  } else {
    domains = operations
      .map((value) => {
        return value.server;
      })
      .filter((value, index, self) => {
        return self.indexOf(value) === index;
      });
    validateDomain(domains, actionName);
  }

  // Need to separate the logic for different flows
  const flow = "flow" in args ? args.flow : "authorizationCode";
  const authInfoArray = operations
    .map((value) => {
      let authInfo;
      switch (flow) {
        case "authorizationCode":
        default:
          authInfo = (value.auth?.authScheme as OpenAPIV3.OAuth2SecurityScheme).flows
            .authorizationCode;
      }

      return {
        authorizationUrl: args.authorizationUrl
          ? args.authorizationUrl
          : authInfo!.authorizationUrl,
        tokenUrl: args.tokenUrl ? args.tokenUrl : authInfo!.tokenUrl,
        refreshUrl: args.refreshUrl ? args.refreshUrl : authInfo!.refreshUrl,
        scopes: args.scope ? parseScopes(args.scope) : Object.keys(authInfo!.scopes),
      };
    })
    .reduce((accumulator: AuthInfo[], currentValue) => {
      if (!accumulator.find((item) => isEqual(item, currentValue))) {
        accumulator.push(currentValue);
      }
      return accumulator;
    }, []);

  if (authInfoArray.length !== 1) {
    throw new OauthAuthInfoInvalid(actionName);
  }
  const authInfo = authInfoArray[0];
  return {
    domain: domains,
    authorizationEndpoint: authInfo.authorizationUrl,
    tokenExchangeEndpoint: authInfo.tokenUrl,
    tokenRefreshEndpoint: authInfo.refreshUrl,
    scopes: authInfo.scopes,
  };
}

export function validateSecret(clientSecret: string): boolean {
  if (typeof clientSecret !== "string") {
    return false;
  }

  if (clientSecret.length > maxSecretLength || clientSecret.length < minSecretLength) {
    return false;
  }

  return true;
}

function validateDomain(domain: string[], actionName: string): void {
  if (domain.length > maxDomainPerOauth) {
    throw new OauthDomainInvalidError(actionName);
  }

  if (domain.length === 0 || domain.includes("")) {
    throw new OauthFailedToGetDomainError(actionName);
  }
}

export function validateUrl(baseUrl: string): boolean {
  try {
    const url = new URL(baseUrl);
    return url.protocol === "https:";
  } catch (error) {
    return false;
  }
}

export function parseScopes(scopes: string | undefined): string[] | undefined {
  if (!scopes) {
    return undefined;
  }
  return scopes.split(",").map((scope) => scope.trim());
}
