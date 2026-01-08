// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Test setup file that runs before all tests.
 * This file mocks native modules that are not needed for e2e tests
 * and may cause issues on certain platforms (e.g., Linux CI).
 */

// Mock keytar to prevent native module loading issues on Linux
// keytar is used by @azure/msal-node-extensions for secure credential storage,
// but is not needed in e2e tests where we mock authentication
const Module = require("module");
const originalRequire = Module.prototype.require;

Module.prototype.require = function (id: string, ...args: any[]) {
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
