// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import { getAbsolutePath } from "../../../utils/common";
import { DriverContext } from "../../interface/commonArgs";
import { CreateApiKeyArgs } from "../interface/createApiKeyArgs";
import { UpdateApiKeyArgs } from "../interface/updateApiKeyArgs";
import { maxDomainPerApiKey } from "./constants";
import { ApiKeyDomainInvalidError } from "../error/apiKeyDomainInvalid";
import { ApiKeyFailedToGetDomainError } from "../error/apiKeyFailedToGetDomain";
import { ApiKeyAuthMissingInSpecError } from "../error/apiKeyAuthMissingInSpec";

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
  context: DriverContext,
  actionName: string
): Promise<string[]> {
  const absolutePath = getAbsolutePath(args.apiSpecPath, context.projectPath);
  const parser = new SpecParser(absolutePath, {
    allowBearerTokenAuth: true, // Currently, API key auth support is actually bearer token auth
    allowMultipleParameters: true,
    projectType: ProjectType.Copilot,
  });
  const listResult = await parser.list();
  const operations = listResult.APIs;

  const filteredOperations = operations.filter((value) => {
    const auth = value.auth;
    return (
      auth &&
      auth.name === args.name &&
      ((auth.authScheme.type === "http" && auth.authScheme.scheme === "bearer") ||
        (auth.authScheme.type === "apiKey" && auth.authScheme.in !== "cookie"))
    );
  });

  if (filteredOperations.length === 0) {
    throw new ApiKeyAuthMissingInSpecError(actionName, args.name);
  }

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
