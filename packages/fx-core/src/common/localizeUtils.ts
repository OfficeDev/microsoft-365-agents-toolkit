// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import * as path from "path";
import * as util from "util";
import { getResourceFolder } from "../folder";
import { Locale } from "./globalVars";

const LocaleStringMap = new Map<string, any>();

function loadTranslationFile(translationFolder: string, locale: string) {
  const nlsFileName = locale ? `package.nls.${locale}.json` : "package.nls.json";
  let nlsFilePath = path.join(translationFolder, nlsFileName);
  if (!fs.pathExistsSync(nlsFilePath)) {
    // if nls file does not exist, just read the default one
    nlsFilePath = path.join(translationFolder, "package.nls.json");
  }
  const json = fs.readJSONSync(nlsFilePath);
  if (json) {
    const existing = LocaleStringMap.get(locale);
    if (existing) {
      LocaleStringMap.set(locale, { ...existing, ...json });
    } else {
      LocaleStringMap.set(locale, json);
    }
  }
}

function getLocaleJson(locale?: string): any {
  locale = locale || "";
  const jsonInMap = LocaleStringMap.get(locale);
  if (jsonInMap) return jsonInMap;

  // fx-core translation files
  loadTranslationFile(getResourceFolder(), locale);
  // template translation files
  loadTranslationFile(path.join(getResourceFolder(), "templates"), locale);
  return LocaleStringMap.get(locale);
}

export function getLocalizedString(key: string, ...params: any[]): string {
  const json = getLocaleJson(Locale);
  let value = json[key];
  if (value && params && params.length > 0) {
    value = util.format(value, ...params);
  }

  if (!value) {
    return getDefaultString(key, ...params);
  }
  return value || "";
}

export function getDefaultString(key: string, ...params: any[]): string {
  const json = getLocaleJson("");
  let value = json[key];
  if (value && params && params.length > 0) {
    value = util.format(value, ...params);
  }
  return value || "";
}
