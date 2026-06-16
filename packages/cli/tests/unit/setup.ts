// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Test setup file that runs before all tests.
 * This file mocks native modules that are not needed for unit tests
 * and may cause issues on certain platforms (e.g., Linux CI).
 */

// Mock keytar to prevent native module loading issues on Linux
// keytar is used by @azure/msal-node-extensions for secure credential storage,
// but is not needed in unit tests where we mock authentication
const Module = require("module");
const originalRequire = Module.prototype.require;

// Keep legacy `require("../../src/foo")` calls working in tests.
require("ts-node/register/transpile-only");

// Mocha-style aliases used by existing tests.
global.before = global.beforeAll;
global.after = global.afterAll;
global.context = global.describe;

Module.prototype.require = function (id, ...args) {
  if (id === "keytar") {
    return {
      getPassword: async () => null,
      setPassword: async () => {},
      deletePassword: async () => true,
      findPassword: async () => null,
      findCredentials: async () => [],
    };
  }
  return originalRequire.apply(this, [id, ...args]);
};

const originalConsoleWarn = console.warn.bind(console);
const originalConsoleError = console.error.bind(console);

function isNoisyTelemetryWarning(args: unknown[]): boolean {
  return args.some(
    (arg) =>
      typeof arg === "string" &&
      arg.includes("ApplicationInsights:An invalid instrumentation key was provided.")
  );
}

console.warn = (...args: unknown[]) => {
  if (isNoisyTelemetryWarning(args)) {
    return;
  }
  originalConsoleWarn(...args);
};

console.error = (...args: unknown[]) => {
  if (isNoisyTelemetryWarning(args)) {
    return;
  }
  originalConsoleError(...args);
};
