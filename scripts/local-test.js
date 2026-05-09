/**
 * local-test.js – Run ATK UI tests locally against the built ATK VSCode extension.
 *
 * Usage (from atk-pipeline repo root):
 *   node scripts/local-test.js [TEST_FILE] [ATK_EXT_PATH]
 *
 * Examples:
 *   node scripts/local-test.js
 *   node scripts/local-test.js teams-bot-create-template
 *   ATK_EXT_PATH=C:\...\packages\vscode-extension node scripts/local-test.js
 *
 * Prerequisites (build ATK extension first):
 *   cd C:\Users\quke\source\atk\microsoft-365-agents-toolkit
 *   pnpm run setup
 *   pnpm build
 */

'use strict';

const path = require('path');
const fs = require('fs');
const { ExTester, ReleaseQuality } = require('./packages/tests/node_modules/vscode-extension-tester');

// ── Config ────────────────────────────────────────────────────────────────────
const DEFAULT_ATK_EXT = 'C:\\Users\\quke\\source\\atk\\microsoft-365-agents-toolkit\\packages\\vscode-extension';

const TEST_FILE    = process.argv[2] || process.env.TEST_FILE    || 'teams-bot-create-template';
const EXT_DEV_PATH = process.argv[3] || process.env.ATK_EXT_PATH || DEFAULT_ATK_EXT;
const OUTPUT_DIR   = process.env.TEST_OUTPUT_DIR || path.join(__dirname, 'test-output');
const STORAGE      = path.join(__dirname, '.vscode-test-storage');
const SPEC         = path.join(__dirname, 'packages', 'tests', 'src', 'ui-test',
                               'copilot-driven', `${TEST_FILE}.test.ts`);
const MOCHARC      = path.join(__dirname, 'packages', 'tests', '.mocharc.js');

// ── Validate ──────────────────────────────────────────────────────────────────
if (!fs.existsSync(SPEC)) {
  console.error(`\n❌ Test spec not found:\n   ${SPEC}`);
  process.exit(1);
}
if (!fs.existsSync(EXT_DEV_PATH)) {
  console.error(`\n❌ ATK extension path not found:\n   ${EXT_DEV_PATH}`);
  console.error('\nBuild the extension first:');
  console.error('  cd C:\\Users\\quke\\source\\atk\\microsoft-365-agents-toolkit');
  console.error('  pnpm run setup && pnpm build\n');
  process.exit(1);
}

fs.mkdirSync(path.join(OUTPUT_DIR, 'screenshots'), { recursive: true });
fs.mkdirSync(path.join(OUTPUT_DIR, 'projects'),    { recursive: true });
process.env.TEST_OUTPUT_DIR = OUTPUT_DIR;

console.log('\n======================================================');
console.log('  ATK Local Test Runner');
console.log(`  Test : ${TEST_FILE}`);
console.log(`  Ext  : ${EXT_DEV_PATH}`);
console.log(`  Out  : ${OUTPUT_DIR}`);
console.log('======================================================\n');

// ── ExTester Strategy ─────────────────────────────────────────────────────────
// ExTester v8 behaviour:
//   coverage=true  → skips installVsix (we don't want to package/install from CWD)
//                  → sets EXTENSION_DEV_PATH = process.cwd() in runTests()
// We monkey-patch process.cwd() temporarily so our ext path propagates correctly.

const tester = new ExTester(STORAGE, ReleaseQuality.Stable, undefined, true /* coverage */);

(async () => {
  try {
    // Step 1: Download VSCode + ChromeDriver (skip extension install because coverage=true)
    console.log('⬇  Downloading VSCode + ChromeDriver (skip if already cached)...');
    await tester.setupRequirements(
      { vscodeVersion: 'latest', installDependencies: false },
      false /* offline */,
      false /* cleanup */
    );

    // Step 2: Run tests — temporarily make process.cwd() return our extension path
    // so ExTester sets EXTENSION_DEV_PATH correctly (it does: process.env.EXTENSION_DEV_PATH = coverage ? process.cwd() : undefined)
    console.log('\n▶  Running test suite...\n');
    const realCwd = process.cwd.bind(process);
    process.cwd = () => EXT_DEV_PATH;

    const code = await tester.runTests(SPEC, {
      vscodeVersion: 'latest',
      config: fs.existsSync(MOCHARC) ? MOCHARC : undefined,
    });

    process.cwd = realCwd;

    // ── Summary ──────────────────────────────────────────────────────────────
    console.log('\n======================================================');
    if (code === 0) {
      console.log('  ✅  All tests PASSED');
    } else {
      console.log(`  ❌  Tests FAILED (exit code: ${code})`);
    }
    console.log(`  Screenshots : ${path.join(OUTPUT_DIR, 'screenshots')}`);
    console.log(`  Results     : ${path.join(OUTPUT_DIR, 'results.json')}`);
    console.log('======================================================\n');

    process.exit(code);
  } catch (err) {
    console.error('\n❌ Fatal error:', err.message);
    process.exit(1);
  }
})();