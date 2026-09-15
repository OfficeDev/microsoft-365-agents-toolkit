// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import { WhitelistFn } from "../../expression/evaluateExpression";

/** Whitelist fn `safeProjectNameLowerCase(appName)` for package names. */
export const safeProjectNameLowerCase: WhitelistFn = (appName: string): string =>
  appName.replace(/[^0-9a-zA-Z]/g, "").toLowerCase();

/** Whitelist fn `pathDelimiter()` for PATH-like launch configuration values. */
export const pathDelimiter: WhitelistFn = (): string => path.delimiter;

/** Parse a comma-joined multiSelect answer into its selected ids. */
export function parseCsv(csv: string): string[] {
  return csv
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
}

/** Whitelist fn `contains(csv, item)` — `"true"` when a multiSelect answer selects `item`. */
export const contains: WhitelistFn = (csv: string, item: string): string =>
  parseCsv(csv).includes(item) ? "true" : "";
