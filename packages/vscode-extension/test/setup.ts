// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
"use strict";

// Ensure tests run with English locale regardless of OS or Node.js navigator.language
if (!process.env.VSCODE_NLS_CONFIG) {
  process.env.VSCODE_NLS_CONFIG = JSON.stringify({ locale: "en" });
}

import * as child_process from "child_process";
import * as os from "os";
import { afterAll, beforeAll, describe, vi } from "vitest";
import { ExtTelemetry } from "../src/telemetry/extTelemetry";
if (os.platform() === "win32") {
  const proc = child_process.spawn("C:\\Windows\\System32\\Reg.exe", ["/?"]);
  proc.on("error", () => {
    console.error("error during reg.exe");
  });
}

if ((Reflect as any).metadata === undefined) {
  require("reflect-metadata");
}

import { initialize } from "./mocks/vscode-mock";

initialize();

// Keep a single vscode module instance so Sinon stubs in tests and runtime code target the same object.
vi.mock("vscode", async () => {
  const mod = await import("./mocks/vscode-mock");
  return mod.mockedVSCode as any;
});

// Redirect globalState to a unique per-worker temp directory so parallel Vitest workers
// don't contend for the same proper-lockfile file lock (which causes 1-second retry delays).
import _setupOs from "os";
import _setupPath from "path";
const _workerStateDir = _setupPath.join(_setupOs.tmpdir(), `teamsfx-test-${process.pid}`);
process.env.USERPROFILE = _workerStateDir; // Windows: os.homedir() reads USERPROFILE
process.env.HOME = _workerStateDir; // Unix: os.homedir() reads HOME

// Some suites call ExtTelemetry APIs directly without creating a Reporter instance.
// Provide a safe default no-op reporter for test runtime.
ExtTelemetry.reporter = {
  addSharedProperty: () => {},
  sendTelemetryEvent: () => {},
  sendTelemetryErrorEvent: () => {},
  sendTelemetryException: () => {},
  dispose: async () => {},
} as any;

// Keep legacy aliases used by existing test files.
Object.assign(globalThis, {
  before: beforeAll,
  after: afterAll,
  context: describe,
});
