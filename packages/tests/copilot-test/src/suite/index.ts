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

  // Optional: run only a specific test file (basename without .js suffix, or full relative path)
  const testFileFilter = process.env.TEST_FILE;

  // glob v7 API: sync returns string[]
  let files = glob.sync("**/*.test.js", { cwd: testsRoot });

  if (testFileFilter) {
    const filterBase = testFileFilter
      .replace(/\.ts$/, ".js")
      .replace(/\.js$/, "");
    files = files.filter((f) => {
      const base = path.basename(f, ".js");
      return base === filterBase || f.includes(filterBase);
    });
    if (files.length === 0) {
      console.warn(
        `[suite/index] TEST_FILE="${testFileFilter}" matched no files — running all tests`,
      );
      files = glob.sync("**/*.test.js", { cwd: testsRoot });
    }
  }

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
