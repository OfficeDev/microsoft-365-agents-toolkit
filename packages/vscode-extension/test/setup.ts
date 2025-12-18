// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
"use strict";

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
