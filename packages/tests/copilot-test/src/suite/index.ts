// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * suite/index.ts - Mocha suite entry, runs inside VSCode extension host
 */
import * as path from "path";
import * as fs from "fs";
import * as os from "os";
import Mocha from "mocha";
import * as glob from "glob";

export async function run(): Promise<void> {
  // One-shot lock: when ATK opens the scaffolded project folder, VS Code reloads
  // the extension host and runs this suite again in the new window. The lock
  // prevents duplicate test runs — only the first host proceeds.
  const lockFile = path.join(
    process.env.TEST_OUTPUT_DIR || os.tmpdir(),
    "atk-copilot-test.lock",
  );
  if (fs.existsSync(lockFile)) {
    console.log(
      "[suite/index] Lock file exists — standby mode (primary host running, never exit)",
    );
    // Never resolve: if this host returns (code 0), VS Code's test runner calls
    // process.exit(0), which kills VS Code and terminates the primary host mid-test.
    // Hanging here keeps VS Code alive until the primary host finishes and VS Code
    // exits on its own.
    await new Promise<void>(() => {}); // intentionally never resolves
    return;
  }
  fs.writeFileSync(lockFile, String(process.pid), "utf8");
  // Clean up lock when the extension host exits so re-runs work.
  process.on("exit", () => {
    try {
      fs.unlinkSync(lockFile);
    } catch {}
  });

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
