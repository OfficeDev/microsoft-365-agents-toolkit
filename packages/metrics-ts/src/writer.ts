// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { join } from "path";
import { tracePoint } from "./rawData";
import { appendFile, ensureDir, ensureDirSync, ensureFile, ensureFileSync } from "fs-extra";
import { EOL, tmpdir } from "os";
import { appendFileSync } from "fs";

const metricsFolder = join(tmpdir(), ".metrics");
const metricsFile = join(metricsFolder, "output.txt");

export async function appendOutput(data: tracePoint): Promise<void> {
  await ensureDir(metricsFolder);
  await ensureFile(metricsFile);
  await appendFile(metricsFile, JSON.stringify(data) + EOL);
  return;
}

export function appendOutputSync(data: tracePoint): void {
  ensureDirSync(metricsFolder);
  ensureFileSync(metricsFile);
  appendFileSync(metricsFile, JSON.stringify(data) + EOL);
  return;
}
