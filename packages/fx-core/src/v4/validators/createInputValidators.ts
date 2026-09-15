// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { capabilityDeclarations } from "../capabilities/declarations";
import { Validator } from "../collectInputs/collectInputs";
import { getLocalizedString } from "../../common/localizeUtils";
import { isValidHttpUrl } from "../../common/stringUtils";
import {
  mcpEntraClientIdRequiredValidator,
  mcpOauthClientIdRequiredValidator,
  mcpOauthClientSecretRequiredValidator,
} from "./mcpCredentialValidators";
import { mcpServerUrlValidator } from "./mcpServerUrlValidator";

export const uriValidator: Validator = (value: string): string | undefined => {
  try {
    new URL(value);
    return undefined;
  } catch {
    return "must be a valid URI";
  }
};

export const openApiUrlValidator: Validator = (value: string): string | undefined => {
  return isValidHttpUrl(value.trim())
    ? undefined
    : getLocalizedString("core.createProjectQuestion.invalidUrl.message");
};

export const graphConnectorNameValidator: Validator = (value: string): string | undefined => {
  return value.trim().length > 0 ? undefined : "must not be empty";
};

export const graphConnectorConnectionIdValidator: Validator = (
  value: string
): string | undefined => {
  const trimmed = value.trim();
  if (trimmed.length < 3) {
    return "must be at least 3 characters";
  }
  if (trimmed.length > 32) {
    return "must be at most 32 characters";
  }
  if (!/^[a-zA-Z0-9]+$/.test(trimmed)) {
    return "must contain only alphanumeric characters";
  }
  const reservedPrefixes = [
    "Microsoft",
    "None",
    "Directory",
    "Exchange",
    "ExchangeArchive",
    "LinkedIn",
    "Mailbox",
    "OneDriveBusiness",
    "SharePoint",
    "Teams",
    "Yammer",
    "Connectors",
    "TaskFabric",
    "PowerBI",
    "Assistant",
    "TopicEngine",
    "MSFT_All_Connectors",
  ];
  const matchedPrefix = reservedPrefixes.find((prefix) =>
    trimmed.toLowerCase().startsWith(prefix.toLowerCase())
  );
  return matchedPrefix === undefined ? undefined : `must not begin with '${matchedPrefix}'`;
};

export function createDefaultCreateInputValidators(): Record<string, Validator> {
  return {
    [capabilityDeclarations.validator.uri.id]: uriValidator,
    [capabilityDeclarations.validator.openApiUrl.id]: openApiUrlValidator,
    [capabilityDeclarations.validator.graphConnectorName.id]: graphConnectorNameValidator,
    [capabilityDeclarations.validator.graphConnectorConnectionId.id]:
      graphConnectorConnectionIdValidator,
    [capabilityDeclarations.validator.mcpOauthClientIdRequired.id]:
      mcpOauthClientIdRequiredValidator,
    [capabilityDeclarations.validator.mcpOauthClientSecretRequired.id]:
      mcpOauthClientSecretRequiredValidator,
    [capabilityDeclarations.validator.mcpEntraClientIdRequired.id]:
      mcpEntraClientIdRequiredValidator,
    [capabilityDeclarations.validator.mcpServerUrl.id]: mcpServerUrlValidator,
  };
}
