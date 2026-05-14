/**
 * local-test.js – Run ATK UI tests with @vscode/test-electron.
 *
 * Uses the official @vscode/test-electron package which launches VSCode
 * with --extensionDevelopmentPath pointing to your locally built ATK extension.
 * Tests run INSIDE the VSCode extension host (can use vscode.* API directly).
 *
 * Usage (from atk-pipeline repo root):
 *   node scripts/local-test.js [TEST_FILE] [ATK_SOURCE_DIR]
 *
 * Examples:
 *   node scripts/local-test.js
 *   node scripts/local-test.js simple-bot-create
 *
 * Prerequisites:
 *   cd C:\Users\quke\source\atk\microsoft-365-agents-toolkit
 *   pnpm run setup && pnpm build
 */
"use strict";

const path = require("path");
const fs = require("fs");
const { spawnSync } = require("child_process");

// ── Config ────────────────────────────────────────────────────────────────────
const DEFAULT_ATK_SRC =
  "C:\\Users\\quke\\source\\atk\\microsoft-365-agents-toolkit";

const TEST_FILE =
  process.argv[2] || process.env.TEST_FILE || "simple-bot-create";
const ATK_SRC = process.argv[3] || process.env.ATK_SOURCE || DEFAULT_ATK_SRC;
const OUTPUT_DIR = path.resolve(
  process.env.TEST_OUTPUT_DIR || path.join(__dirname, "..", "test-output"),
);

const EXT_DEV_PATH = path.join(ATK_SRC, "packages", "vscode-extension");
const ATK_TESTS = path.join(ATK_SRC, "packages", "tests");
const EXTESTER_PATH = path.join(
  ATK_TESTS,
  "node_modules",
  "@vscode",
  "test-electron",
);

// Our runTest.ts (compiled to JS) will be placed in ATK_TESTS out/ dir
const RUNNER_SRC = path.resolve(__dirname, "..", "src", "runTest.ts");
const TEST_SRC = path.resolve(
  __dirname,
  "..",
  "packages",
  "tests",
  "src",
  "ui-test",
  "copilot-driven",
  `${TEST_FILE}.test.ts`,
);

// ── Validate ──────────────────────────────────────────────────────────────────
const errors = [];
if (!fs.existsSync(TEST_SRC))
  errors.push("Test spec not found:\n   " + TEST_SRC);
if (!fs.existsSync(EXT_DEV_PATH))
  errors.push(
    "ATK extension not built:\n   " +
      EXT_DEV_PATH +
      "\n   Run: pnpm run setup && pnpm build",
  );

// Check for @vscode/test-electron in ATK workspace OR install it
const testElectronGlobal = path.join(
  require("os").homedir(),
  ".npm-global",
  "node_modules",
  "@vscode",
  "test-electron",
);
const hasTestElectron =
  fs.existsSync(EXTESTER_PATH) || fs.existsSync(testElectronGlobal);
if (!hasTestElectron) {
  console.log(
    "\u2139 @vscode/test-electron not found in ATK workspace, installing...",
  );
  const install = spawnSync(
    "npm",
    [
      "install",
      "-g",
      "@vscode/test-electron",
      "@vscode/test-cli",
      "ts-node",
      "mocha",
      "glob",
    ],
    {
      stdio: "inherit",
      shell: true,
    },
  );
  if (install.status !== 0) {
    errors.push("@vscode/test-electron install failed");
  }
}

if (errors.length) {
  errors.forEach((e) => console.error("\n\u274c", e));
  process.exit(1);
}

// ── Copy test files into ATK workspace ───────────────────────────────────────
const DST_COPILOT = path.join(ATK_TESTS, "src", "ui-test", "copilot-driven");
const DST_SUITE = path.join(DST_COPILOT, "suite");
fs.mkdirSync(DST_SUITE, { recursive: true });

// Copy all our test files
const srcDir = path.resolve(__dirname, "..", "src");
[
  ["runTest.ts", path.join(DST_COPILOT, "runTest.ts")],
  ["suite/index.ts", path.join(DST_SUITE, "index.ts")],
  [`${TEST_FILE}.test.ts`, path.join(DST_COPILOT, `${TEST_FILE}.test.ts`)],
].forEach(([rel, dst]) => {
  const src = path.join(srcDir, rel);
  if (fs.existsSync(src)) {
    fs.copyFileSync(src, dst);
    console.log("\u2714 Copied", rel, "\u2192", dst);
  }
});

// ── Prepare output dir ────────────────────────────────────────────────────────
["screenshots", "projects"].forEach((d) =>
  fs.mkdirSync(path.join(OUTPUT_DIR, d), { recursive: true }),
);

console.log("\n======================================================");
console.log("  ATK Local Test Runner (@vscode/test-electron)");
console.log("  Test    :", TEST_FILE);
console.log("  Ext     :", EXT_DEV_PATH);
console.log("  ATK src :", ATK_SRC);
console.log("  Output  :", OUTPUT_DIR);
console.log("======================================================\n");

// ── Compile & run with ts-node ────────────────────────────────────────────────
// We run runTest.ts directly via ts-node (available in ATK workspace)
const tsNode = path.join(
  ATK_TESTS,
  "node_modules",
  ".bin",
  process.platform === "win32" ? "ts-node.cmd" : "ts-node",
);
const tsNodeBin = fs.existsSync(tsNode) ? tsNode : "ts-node";
const runnerFile = path.join(DST_COPILOT, "runTest.ts");

console.log("\u25b6 Compiling and running with ts-node...\n");

const result = spawnSync(
  tsNodeBin,
  ["--project", path.join(ATK_TESTS, "tsconfig.json"), runnerFile],
  {
    cwd: ATK_TESTS,
    stdio: "inherit",
    shell: false,
    env: {
      ...process.env,
      ATK_EXT_PATH: EXT_DEV_PATH,
      TEST_OUTPUT_DIR: OUTPUT_DIR,
      TEST_FILE: TEST_FILE,
    },
  },
);

const code = result.status ?? 1;
console.log("\n======================================================");
console.log(
  code === 0
    ? "  \u2705  Tests PASSED"
    : "  \u274c  Tests FAILED (exit: " + code + ")",
);
console.log("  Screenshots :", path.join(OUTPUT_DIR, "screenshots"));
console.log("  Results     :", path.join(OUTPUT_DIR, "results.json"));
console.log("======================================================\n");
process.exit(code);
