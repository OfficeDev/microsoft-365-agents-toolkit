// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * suite/index.ts – Mocha suite entry, runs inside VSCode extension host
 */
import * as path from "path";
import Mocha from "mocha";
import { glob } from "glob";

export async function run(): Promise<void> {
  const mocha = new Mocha({
    ui: "bdd",
    timeout: 5 * 60 * 1000,
    color: true,
    reporter: "spec",
  });

  const testsRoot = path.resolve(__dirname, "..");

  const files = await glob("**/*.test.js", { cwd: testsRoot });
  files.forEach((f) => mocha.addFile(path.resolve(testsRoot, f)));

  return new Promise((resolve, reject) => {
    mocha.run((failures) => {
      if (failures > 0) {
        reject(new Error(`${failures} test(s) failed`));
      } else {
        resolve();
      }
    });
  });
}