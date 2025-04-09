// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { getAbsolutePath } from "../../../utils/common";
import { CreateApiKeyArgs } from "../interface/createApiKeyArgs";
import { UpdateApiKeyArgs } from "../interface/updateApiKeyArgs";
import { maxDomainPerApiKey, telemetryKeys } from "./constants";
import { ApiKeyDomainInvalidError } from "../error/apiKeyDomainInvalid";
import { ApiKeyFailedToGetDomainError } from "../error/apiKeyFailedToGetDomain";
import { ApiKeyAuthMissingInSpecError } from "../error/apiKeyAuthMissingInSpec";
import { WrapDriverContext } from "../../util/wrapUtil";
import { listAPIInfo } from "../../../../common/daSpecParser";

// Needs to validate the parameters outside of the function
export function loadStateFromEnv(
  outputEnvVarNames: Map<string, string>
): Record<string, string | undefined> {
  const result: Record<string, string | undefined> = {};
  for (const [propertyName, envVarName] of outputEnvVarNames) {
    result[propertyName] = process.env[envVarName];
  }
  return result;
}

// TODO: need to add logic to read domain from env if need to support non-lifecycle commands
export async function getDomain(
  args: CreateApiKeyArgs | UpdateApiKeyArgs,
  context: WrapDriverContext,
  actionName: string
): Promise<string[]> {
  const absolutePath = getAbsolutePath(args.apiSpecPath!, context.projectPath);

  const listResult = await listAPIInfo(absolutePath);
  const operations = listResult.APIs;

  const filteredOperations = operations.filter((value) => {
    const auth = value.auth;
    return (
      auth &&
      auth.name === args.name &&
      ((auth.authScheme.type === "http" && auth.authScheme.scheme.toLowerCase() === "bearer") ||
        (auth.authScheme.type === "apiKey" && auth.authScheme.in !== "cookie"))
    );
  });

  if (filteredOperations.length === 0) {
    throw new ApiKeyAuthMissingInSpecError(actionName, args.name);
  }

  const isCustomAPIKey =
    filteredOperations[0].auth!.authScheme.type === "apiKey" ? "true" : "false";

  context.addTelemetryProperties({ [telemetryKeys.isCustomAPIKey]: isCustomAPIKey });

  const servers = filteredOperations.map((value) => value.server);

  const uniqueServerUrls = servers.filter((value, index, self) => self.indexOf(value) === index);

  return uniqueServerUrls;
}

export function validateDomain(domain: string[], actionName: string): void {
  if (domain.length > maxDomainPerApiKey) {
    throw new ApiKeyDomainInvalidError(actionName);
  }

  if (domain.length === 0 || domain.includes("")) {
    throw new ApiKeyFailedToGetDomainError(actionName);
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
