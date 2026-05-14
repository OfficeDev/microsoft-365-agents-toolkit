// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * suite/index.ts - Mocha suite entry, runs inside VSCode extension host
 */
import * as path from "path";
import Mocha from "mocha";
import * as glob from "glob";

export async function run(): Promise<void> {
  // Use "tdd" UI so suite()/test() work (ATK convention)
  const mocha = new Mocha({
    ui: "tdd",
    timeout: 5 * 60 * 1000,
    color: true,
    reporter: "spec",
  });

  const testsRoot = path.resolve(__dirname, "..");

  // glob v7 API: sync returns string[]
  const files = glob.sync("**/*.test.js", { cwd: testsRoot });
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
