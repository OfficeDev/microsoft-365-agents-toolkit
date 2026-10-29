const assert = require("node:assert/strict");
const { test } = require("node:test");
const fs = require("node:fs");
const path = require("node:path");

const workflowPath = path.resolve(
  __dirname,
  "..",
  "..",
  "workflows",
  "eval-microsoft-365-agents-toolkit.yml",
);
const validationWorkflowPath = path.resolve(
  __dirname,
  "..",
  "..",
  "workflows",
  "validate-microsoft-365-agents-toolkit-evals.yml",
);
const setupPath = path.resolve(
  __dirname,
  "..",
  "..",
  "..",
  "evals",
  "microsoft-365-agents-toolkit",
  "setup.sh",
);

const workflow = fs.readFileSync(workflowPath, "utf8").replaceAll("\r\n", "\n");
const validationWorkflow = fs
  .readFileSync(validationWorkflowPath, "utf8")
  .replaceAll("\r\n", "\n");
const evaluateJob = workflow.slice(workflow.indexOf("\n  evaluate:\n"));
const setup = fs.readFileSync(setupPath, "utf8");

test("untrusted pull requests cannot run the credentialed skill evaluation", () => {
  assert.doesNotMatch(workflow, /\n  pull_request:\n/);
  assert.doesNotMatch(workflow, /pull_request_target/);
  assert.doesNotMatch(evaluateJob, /github\.event\.pull_request\.(?:head|base)/);
  assert.doesNotMatch(evaluateJob, /Checkout PR skill|Overlay PR skill/);
  assert.match(validationWorkflow, /\n  pull_request:\n/);
  assert.doesNotMatch(validationWorkflow, /pull_request_target|environment:|secrets\./);
});

test("non-deployment evaluation exposes only the Copilot credential", () => {
  assert.match(evaluateJob, /GITHUB_TOKEN: \$\{\{ secrets\.COPILOT_GITHUB_TOKEN \}\}/);
  assert.doesNotMatch(evaluateJob, /M365_ACCOUNT_|AZURE_ACCOUNT_|AZURE_SUBSCRIPTION_ID/);
  assert.doesNotMatch(setup, /M365_ACCOUNT_|AZURE_ACCOUNT_|AZURE_SUBSCRIPTION_ID/);
});
