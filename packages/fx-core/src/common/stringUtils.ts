// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as crypto from "crypto";
import * as Handlebars from "handlebars";
import { URL } from "url";
import * as uuid from "uuid";
import { FailedToParseResourceIdError } from "../error/common";
import { getLocalizedString } from "./localizeUtils";
import { secretMasker } from "./secretmasker/masker";

const SECRET_REPLACE = "<REDACTED:secret>";

export interface MaskSecretOptions {
  threshold?: number;
  whiteList?: string[];
  replace?: string;
}

export function maskSecret(inputText?: string, option?: MaskSecretOptions): string {
  if (!inputText) return "";
  const replace = option?.replace || SECRET_REPLACE;
  let output = maskSecretFromEnv(inputText);
  output = secretMasker.maskSecret(output, replace, option?.whiteList);
  return output;
}

export function maskSecretFromEnv(stdout: string, replace = SECRET_REPLACE): string {
  for (const key of Object.keys(process.env)) {
    if (key.startsWith("SECRET_")) {
      const value = process.env[key];
      if (value) {
        const escapedValue = value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        stdout = stdout.replace(new RegExp(escapedValue, "g"), replace);
      }
    }
  }
  return stdout;
}

export function convertToAlphanumericOnly(appName: string): string {
  return appName.replace(/[^\da-zA-Z]/g, "");
}

Handlebars.registerHelper("contains", (value, array) => {
  array = array instanceof Array ? array : [array];
  return array.indexOf(value) > -1 ? this : "";
});
Handlebars.registerHelper("notContains", (value, array) => {
  array = array instanceof Array ? array : [array];
  return array.indexOf(value) == -1 ? this : "";
});
Handlebars.registerHelper("equals", (value, target) => {
  return value === target ? this : "";
});

export function getResourceGroupNameFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(/\/resourceGroups\/([^\/]*)\//i, resourceId);
  if (!result) {
    throw new FailedToParseResourceIdError("resource group name", resourceId);
  }
  return result;
}

export function parseFromResourceId(pattern: RegExp, resourceId: string): string {
  const result = resourceId.match(pattern);
  return result ? result[1].trim() : "";
}

export function getUuid(): string {
  return uuid.v4();
}

export function getHashedEnv(envName: string): string {
  return crypto.createHash("sha256").update(envName).digest("hex");
}

export function loadingOptionsPlaceholder(): string {
  return getLocalizedString("ui.select.LoadingOptionsPlaceholder");
}

export function loadingDefaultPlaceholder(): string {
  return getLocalizedString("ui.select.LoadingDefaultPlaceholder");
}

export function isValidHttpUrl(input: string): boolean {
  let url;
  try {
    url = new URL(input);
    return url.protocol === "http:" || url.protocol === "https:";
  } catch (e) {
    return false;
  }
}
