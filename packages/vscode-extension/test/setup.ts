// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
"use strict";

// Mock keytar to prevent native module loading issues on Linux
// keytar is used by @azure/msal-node-extensions for secure credential storage,
// but is not needed in unit tests where we mock authentication
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

import * as child_process from "child_process";
import * as os from "os";
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
