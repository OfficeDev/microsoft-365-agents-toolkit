// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * runTest.ts – entry point for @vscode/test-electron
 *
 * Launched via: ts-node runTest.ts
 * Downloads VSCode, then launches it with --extensionDevelopmentPath pointing
 * at the locally built ATK extension. Tests run INSIDE the VSCode extension host.
 */
import * as path from "path";
import { runTests, downloadAndUnzipVSCode } from "@vscode/test-electron";
import * as fs from "fs";

async function main() {
  const extensionDevelopmentPath =
    process.env.ATK_EXT_PATH ||
    path.resolve(__dirname, "../../../../../../vscode-extension");

  const outputDir =
    process.env.TEST_OUTPUT_DIR ||
    path.resolve(__dirname, "../../../../../test-output");

  // The compiled test suite entry point – ts-node compiles on the fly
  // For @vscode/test-electron we need the compiled .js path.
  // We compile runTest.ts manually with ts-node, but extensionTestsPath
  // must be a compiled .js module that VSCode can require inside the host.
  // Compile the suite to a temp dir first.
  const tmpOut = path.join(outputDir, "compiled-tests");
  fs.mkdirSync(tmpOut, { recursive: true });

  console.log("=== @vscode/test-electron Runner ===");
  console.log("extensionDevelopmentPath:", extensionDevelopmentPath);
  console.log("outputDir:", outputDir);

  if (!fs.existsSync(extensionDevelopmentPath)) {
    console.error("ERROR: ATK extension not found:", extensionDevelopmentPath);
    console.error("Build it first: pnpm run setup && pnpm build");
    process.exit(1);
  }

  // Compile our test suite with tsc so VSCode's extension host can require it
  const { spawnSync } = require("child_process");
  const tscBin = path.join(__dirname, "../../../node_modules/.bin/tsc");
  const tsc = fs.existsSync(tscBin) ? tscBin : "tsc";

  // Create a temporary tsconfig pointing to our suite
  const tmpTsConfig = path.join(tmpOut, "tsconfig.json");
  fs.writeFileSync(tmpTsConfig, JSON.stringify({
    compilerOptions: {
      module: "commonjs",
      target: "ES2020",
      lib: ["ES2020"],
      esModuleInterop: true,
      resolveJsonModule: true,
      strict: false,
      outDir: tmpOut,
      rootDir: path.dirname(__dirname)
    },
    include: [path.join(__dirname, "**/*.ts")],
    exclude: ["node_modules"]
  }, null, 2));

  console.log("Compiling test suite...");
  const compile = spawnSync(tsc, ["--project", tmpTsConfig], {
    stdio: "inherit", shell: true
  });
  if (compile.status !== 0) {
    console.warn("TypeScript compile had errors – attempting to run anyway");
  }

  // The compiled suite index relative to tmpOut
  const suiteDir = path.join(__dirname, "suite");
  const compiledSuiteIndex = path.join(tmpOut, path.relative(path.dirname(__dirname), suiteDir), "index");

  console.log("extensionTestsPath:", compiledSuiteIndex);

  try {
    await runTests({
      extensionDevelopmentPath,
      extensionTestsPath: compiledSuiteIndex,
      launchArgs: [
        "--disable-workspace-trust",
        "--skip-welcome",
        "--skip-release-notes",
        `--user-data-dir=${path.join(outputDir, "vscode-user-data")}`,
        "--no-sandbox",           // needed inside Docker
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