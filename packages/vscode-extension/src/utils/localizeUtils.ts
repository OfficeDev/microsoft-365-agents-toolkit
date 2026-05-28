// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import fs from "fs-extra";
import { context } from "../globalVariables";
import VsCodeLogInstance from "../commonlib/log";

let loadedCollection: Record<string, string> | undefined = undefined;
let defaultCollection: Record<string, string> | undefined = undefined;
let askedForCollection: Record<string, string> = {};
export let loadedLocale: string;

export function localize(key: string, defValue?: string) {
  return getString(key, defValue);
}

export function getDefaultString(key: string, defValue?: string) {
  loadDefaultStrings();
  return getLocalizedString(key, true, defValue);
}

function getString(key: string, defValue?: string) {
  if (shouldReloadLocale()) {
    loadLocalizedStrings();
  }
  return getLocalizedString(key, false, defValue);
}

export function _resetCollections(): void {
  loadedLocale = "";
  defaultCollection = undefined;
  loadedCollection = undefined;
  askedForCollection = {};
}

export function _getAskedForCollection(): Record<string, string> {
  return askedForCollection;
}

function shouldReloadLocale(): boolean {
  return !loadedCollection;
}

declare let navigator: { language: string } | undefined;

export function parseLocale(): string {
  try {
    if (navigator?.language) {
      return navigator.language.toLowerCase();
    }
  } catch {}
  const vscodeLocale = process.env.VSCODE_NLS_CONFIG
    ? (JSON.parse(process.env.VSCODE_NLS_CONFIG) as Record<string, string>).locale
    : undefined;
  VsCodeLogInstance.info(`Current VS Code locale is: ${vscodeLocale ?? ""}`);
  return vscodeLocale ?? "en-us";
}

function getLocalizedString(key: string, isDefault: boolean, defValue?: string): string {
  let collection = defaultCollection;

  if (!isDefault && loadedCollection && loadedCollection.hasOwnProperty(key)) {
    collection = loadedCollection;
  }
  if (collection === undefined) {
    throw new Error(`Localizations haven't been loaded yet for key: ${key}`);
  }
  let result = collection[key];
  if (!result && defValue) {
    result = defValue;
  }
  askedForCollection[key] = result;

  return result;
}

/**
 * Load localized strings according to current locale. By default, load package.nls.json if target locale doesn't exist.
 */
export function loadLocalizedStrings(): void {
  loadedLocale = parseLocale();

  const nlsFile = path.join(
    context ? context.extensionPath : "",
    `package.nls.${loadedLocale}.json`
  );
  if (fs.pathExistsSync(nlsFile)) {
    loadedCollection = fs.readJsonSync(nlsFile) as Record<string, string> | undefined;
  } else {
    if (loadedLocale !== "en" && loadedLocale !== "en-us") {
      VsCodeLogInstance.error(
        `No localized strings file found for locale: ${loadedLocale}, will fallback to default one.`
      );
    }
    loadedCollection = {};
  }

  loadDefaultStrings();
}

function loadDefaultStrings(): void {
  if (!defaultCollection) {
    const defaultNlsFile = path.join(context ? context.extensionPath : "", "package.nls.json");
    if (fs.pathExistsSync(defaultNlsFile)) {
      defaultCollection = fs.readJsonSync(defaultNlsFile) as Record<string, string> | undefined;
    } else {
      defaultCollection = {};
      VsCodeLogInstance.error(`No default localized strings file found.`);
    }
  }
}
