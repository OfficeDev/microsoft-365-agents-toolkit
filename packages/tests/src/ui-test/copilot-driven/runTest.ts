// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * runTest.ts - entry point for @vscode/test-electron
 * Run via: ts-node runTest.ts   (cwd: packages/tests)
 */
import * as path from "path";
import * as fs from "fs";
import * as cp from "child_process";
import { runTests } from "@vscode/test-electron";

// HERE = packages/tests/src/ui-test/copilot-driven
const HERE = __dirname;
// TESTS_ROOT = packages/tests
const TESTS_ROOT = path.resolve(HERE, "../../..");

async function compileSuite(tmpOut: string): Promise<string> {
  fs.mkdirSync(tmpOut, { recursive: true });

  const tsconfigPath = path.join(HERE, "_tsconfig.build.json");
  const tsconfig = {
    compilerOptions: {
      module: "commonjs",
      target: "ES2020",
      lib: ["ES2020"],
      esModuleInterop: true,
      resolveJsonModule: true,
      strict: false,
      skipLibCheck: true,
      outDir: tmpOut,
      rootDir: HERE,
      types: ["vscode", "node", "mocha"],
      typeRoots: [path.join(TESTS_ROOT, "node_modules", "@types")],
    },
    include: ["**/*.ts"],
    exclude: ["node_modules", "_tsconfig.build.json"],
  };
  fs.writeFileSync(tsconfigPath, JSON.stringify(tsconfig, null, 2), "utf8");

  const tscBin = path.join(
    TESTS_ROOT,
    "node_modules",
    ".bin",
    process.platform === "win32" ? "tsc.CMD" : "tsc"
  );

  console.log("Compiling test suite...");
  const result = cp.spawnSync(tscBin, ["--project", tsconfigPath], {
    cwd: HERE,
    stdio: "inherit",
    // .CMD files on Windows require shell:true
    shell: process.platform === "win32",
  });

  fs.rmSync(tsconfigPath, { force: true });

  if (result.status !== 0) {
    console.warn("tsc had errors - continuing anyway");
  }
  return path.join(tmpOut, "suite", "index");
}

async function main() {
  const extensionDevelopmentPath =
    process.env.ATK_EXT_PATH ||
    path.resolve(TESTS_ROOT, "../../packages/vscode-extension");

  const outputDir =
    process.env.TEST_OUTPUT_DIR ||
    path.resolve(TESTS_ROOT, "../../test-output");

  // Put compiled output inside TESTS_ROOT/out/ so the extension host
  // can resolve mocha/glob from packages/tests/node_modules via normal CJS lookup
  const tmpOut = path.join(TESTS_ROOT, "out", "copilot-driven");

  console.log("=== @vscode/test-electron Runner ===");
  console.log("Ext:", extensionDevelopmentPath);
  console.log("Out:", outputDir);

  if (!fs.existsSync(extensionDevelopmentPath)) {
    console.error("ATK extension not found:", extensionDevelopmentPath);
    process.exit(1);
  }

  const extensionTestsPath = await compileSuite(tmpOut);
  console.log("Suite:", extensionTestsPath);

  if (!fs.existsSync(extensionTestsPath + ".js")) {
    console.error("Compiled suite missing:", extensionTestsPath + ".js");
    process.exit(1);
  }

  try {
    await runTests({
      extensionDevelopmentPath,
      extensionTestsPath,
      launchArgs: [
        "--disable-workspace-trust",
        "--skip-welcome",
        "--skip-release-notes",
        // Install ATK extension dependencies so it can activate
        "--install-extension", "redhat.vscode-yaml",
        `--user-data-dir=${path.join(outputDir, "vscode-user-data")}`,
        "--no-sandbox",
      ],
      version: "stable",
      extensionTestsEnv: {
        TEST_OUTPUT_DIR: outputDir,
        TEST_FILE: process.env.TEST_FILE || "teams-bot-create-template",
        ATK_EXT_PATH: extensionDevelopmentPath,
      },
    });
  } catch (err) {
    console.error("Test run failed:", err);
    process.exit(1);
  }
}

main();