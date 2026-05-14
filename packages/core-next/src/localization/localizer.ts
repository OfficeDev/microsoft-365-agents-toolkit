// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import * as fs from "fs";
import * as util from "util";

type StringMap = Record<string, string>;

/**
 * Localizer loads package.nls.json bundles and provides locale-aware string lookup.
 *
 * Reuses the same file format as VS Code / fx-core:
 *   - package.nls.json (English default)
 *   - package.nls.{locale}.json (translated)
 *
 * Fully injectable — no global locale state.
 */
export class Localizer {
  private defaultBundle: StringMap = {};
  private localizedBundle: StringMap = {};
  private readonly locale: string;

  constructor(locale?: string) {
    this.locale = normalizeLocale(locale ?? process.env.VSCODE_NLS_CONFIG);
  }

  /**
   * Load string bundles from a directory.
   * Looks for package.nls.json and package.nls.{locale}.json.
   */
  loadBundle(resourceDir: string): void {
    const defaultPath = path.join(resourceDir, "package.nls.json");
    this.defaultBundle = readJsonSafe(defaultPath);

    if (this.locale) {
      const localePath = path.join(resourceDir, `package.nls.${this.locale}.json`);
      this.localizedBundle = readJsonSafe(localePath);
    }
  }

  /**
   * Get a localized string with positional parameter substitution.
   * Falls back to the English default if no translation exists.
   *
   * @param key - The string key in package.nls.json
   * @param params - Positional parameters for util.format()-style placeholders
   */
  getString(key: string, ...params: unknown[]): string {
    const template = this.localizedBundle[key] ?? this.defaultBundle[key] ?? key;
    return params.length > 0 ? util.format(template, ...params) : template;
  }

  /**
   * Get the English default string (ignores locale).
   */
  getDefaultString(key: string, ...params: unknown[]): string {
    const template = this.defaultBundle[key] ?? key;
    return params.length > 0 ? util.format(template, ...params) : template;
  }
}

/**
 * Factory with sensible defaults.
 */
export function createLocalizer(locale?: string, resourceDir?: string): Localizer {
  const loc = new Localizer(locale);
  if (resourceDir) {
    loc.loadBundle(resourceDir);
  }
  return loc;
}

function normalizeLocale(raw: string | undefined): string {
  if (!raw) return "";
  // VSCODE_NLS_CONFIG is JSON: { "locale": "zh-cn", ... }
  try {
    const parsed = JSON.parse(raw);
    if (typeof parsed === "object" && parsed.locale) {
      return String(parsed.locale).toLowerCase();
    }
  } catch {
    // Raw string locale
  }
  return raw.toLowerCase();
}

function readJsonSafe(filePath: string): StringMap {
  try {
    const content = fs.readFileSync(filePath, "utf-8");
    return JSON.parse(content) as StringMap;
  } catch {
    return {};
  }
}
