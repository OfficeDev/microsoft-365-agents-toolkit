// Selects the 10 most relevant smoke test plans for the current branch by
// asking Copilot SDK (claude-opus-4.6) to diff against `dev` and pick plans
// from packages/tests/vscuse/vscode-test-cases/plans/, then overwrites
// packages/tests/vscuse/smoking-test-cases.json.
//
// Required env vars:
//   COPILOT_GITHUB_TOKEN   - PAT for Copilot SDK
//   REPO_ROOT              - absolute path to repo root
//   CHANGED_FILES_PATH     - path to a file containing changed files (one per line)

import { CopilotClient, approveAll } from "@github/copilot-sdk";
import fs from "fs";
import path from "path";

if (process.stdout._handle?.setBlocking) process.stdout._handle.setBlocking(true);
if (process.stderr._handle?.setBlocking) process.stderr._handle.setBlocking(true);

const repoRoot = process.env.REPO_ROOT;
const changedFilesPath = process.env.CHANGED_FILES_PATH;
const githubToken = process.env.COPILOT_GITHUB_TOKEN;

if (!repoRoot) throw new Error("REPO_ROOT env var is required");
if (!changedFilesPath) throw new Error("CHANGED_FILES_PATH env var is required");

const plansDir = path.join(repoRoot, "packages/tests/vscuse/vscode-test-cases/plans");
const targetFile = path.join(repoRoot, "packages/tests/vscuse/smoking-test-cases.json");

if (!fs.existsSync(plansDir)) throw new Error("Plans directory not found: " + plansDir);
if (!fs.existsSync(targetFile)) throw new Error("Target file not found: " + targetFile);

const allPlans = fs
  .readdirSync(plansDir)
  .filter((f) => f.endsWith(".json"))
  .map((f) => f.replace(/\.json$/, ""));
console.log(`Found ${allPlans.length} candidate test plans in plans/`);

const systemMessage = `You are a test selection assistant for the Microsoft 365 Agents Toolkit repo.

## Environment
- OS: Linux (Ubuntu). Use bash for all shell commands. Do NOT use powershell.
- The repository is checked out at: ${repoRoot}
- You are running in CI; be efficient and avoid unnecessary tool calls.

## Inputs
- List of files changed in this branch vs dev: ${changedFilesPath}
- Directory of all available smoke test plan JSON files: ${plansDir}
  (Each filename without the .json extension is a test case ID. There are ${allPlans.length} candidates.)
- File you must overwrite: ${targetFile}

## Your task
1. Read ${changedFilesPath} to see what changed in this branch.
2. Inspect a small sampling of the most informative changed source files (e.g. under packages/fx-core, packages/vscode-extension, packages/cli, templates/) to understand which capabilities, templates, languages, or debug flows are affected. Do NOT read every changed file — pick a representative subset.
3. List the test plan filenames in ${plansDir} (you can use \`ls\`). Each filename encodes capability + language + debug type, e.g. "Basic_Custom_Engine_Azure_OpenAI_ts_Copilot_Remote_Debug", "Message_extension_ts_remote_debug", "DA_Oauth_js_Remote_Debug", "Teams_Agent_With_Data_AI_Search_Azure_OpenAI_ts_Remote_Debug".
4. Pick exactly 10 test plan IDs whose scope is MOST relevant to the changed files. Match by:
   - Capability keywords: Custom_Engine, Message_extension, DA_/Declarative_Agent, Teams_Agent, Bot, Tab, API_Plugin
   - Language: ts / js / py
   - Debug flow: Local_Debug, Remote_Debug, playground
   - Template/sample names referenced in the diff
5. If changes are broad or unclear, choose a balanced cross-section covering: custom engine agent, declarative agent, message extension, teams agent with data, and a basic bot/tab — across ts/js/py and Local/Remote debug.
6. Overwrite ${targetFile} with EXACTLY this JSON shape (pretty-printed, 4-space indent to match existing style):
   {
       "smokeTestCases": [
           "<id1>",
           ...
           "<id10>"
       ]
   }
   - Exactly 10 entries.
   - Each entry MUST be the basename (without .json) of a real file in ${plansDir}.
   - No duplicates.
7. After writing, cat the file to stdout and briefly explain (1-3 sentences) why these 10 were chosen.

Do NOT modify any other files. Do NOT run tests, builds, git, or network commands beyond what is needed to read the listed inputs.`;

console.log("Starting Copilot SDK client (model: claude-opus-4.6)...");
const client = new CopilotClient({ logLevel: "info", githubToken });
await client.start();

const session = await client.createSession({
  model: "claude-opus-4.6",
  workingDirectory: repoRoot,
  onPermissionRequest: approveAll,
  streaming: true,
  excludedTools: ["powershell"],
  systemMessage: { content: systemMessage },
});

session.on("assistant.message_delta", (e) => process.stdout.write(e.data.deltaContent));
session.on("tool.execution_start", (e) => {
  console.log(`\n[tool:start] ${e.data.toolName}`);
});
session.on("tool.execution_complete", (e) => {
  console.log(`[tool:done] ${e.data.toolCallId} ${e.data.success ? "ok" : "FAIL"}`);
});

console.log("\n" + "=".repeat(60));
await session.sendAndWait(
  {
    prompt:
      "Analyze the branch changes and update packages/tests/vscuse/smoking-test-cases.json with the 10 most relevant smoke test plan IDs, following the rules in your system message.",
  },
  2_147_483_647,
);
console.log("\n" + "=".repeat(60));

await session.destroy();
await client.stop();

// Validate the result
const planSet = new Set(allPlans);
const result = JSON.parse(fs.readFileSync(targetFile, "utf8"));
if (!Array.isArray(result.smokeTestCases)) {
  throw new Error("smokeTestCases is not an array: " + JSON.stringify(result));
}
if (result.smokeTestCases.length !== 10) {
  throw new Error(`Expected exactly 10 smokeTestCases, got ${result.smokeTestCases.length}`);
}
const seen = new Set();
for (const name of result.smokeTestCases) {
  if (typeof name !== "string") throw new Error("Non-string entry: " + name);
  if (seen.has(name)) throw new Error("Duplicate entry: " + name);
  if (!planSet.has(name)) throw new Error(`Test case not found in plans/: ${name}`);
  seen.add(name);
}
console.log("\n✓ Validated 10 unique smoke test cases:");
for (const n of result.smokeTestCases) console.log("  - " + n);
process.exit(0);
