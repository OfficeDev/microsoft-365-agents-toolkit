// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Table from "cli-table3";

export type OutputFormat = "text" | "json";

export interface OutputOptions {
  format: OutputFormat;
}

/**
 * Print structured data as JSON or a human-readable table.
 */
export function printResult(
  data: Record<string, unknown> | Record<string, unknown>[],
  options: OutputOptions
): void {
  if (options.format === "json") {
    console.log(JSON.stringify(data, null, 2));
    return;
  }
  if (Array.isArray(data)) {
    printTable(data);
  } else {
    printKeyValue(data);
  }
}

function printTable(rows: Record<string, unknown>[]): void {
  if (rows.length === 0) {
    console.log("(no results)");
    return;
  }
  const keys = Object.keys(rows[0]);
  const table = new Table({ head: keys });
  for (const row of rows) {
    table.push(keys.map((k) => String(row[k] ?? "")));
  }
  console.log(table.toString());
}

function printKeyValue(obj: Record<string, unknown>): void {
  const table = new Table();
  for (const [key, value] of Object.entries(obj)) {
    table.push({ [key]: String(value ?? "") });
  }
  console.log(table.toString());
}
