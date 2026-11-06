// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const { mkdirSync, mkdtempSync, readdirSync, rmSync } = require("node:fs");
const { tmpdir } = require("node:os");
const path = require("node:path");

const cli = path.resolve(__dirname, "..", "cli.js");
const env = {
  ...process.env,
  TEMPLATE_VERSION: "local",
  CI_ENABLED: "false",
  TEAMSFX_WORKIQ_PUBLISH_TARGET: "workiq",
  FORCE_COLOR: "0",
};
delete env.NO_COLOR;

function run(args, success = true, environment = {}, cwd = process.cwd()) {
  const child = spawnSync(process.execPath, [cli, ...args, "--telemetry", "false"], {
    encoding: "utf8",
    timeout: 60000,
    env: { ...env, ...environment },
    cwd,
  });
  if (child.error) throw child.error;
  const output = child.stdout + child.stderr;
  assert.equal(child.signal, null, output);
  if (success) assert.equal(child.status, 0, output);
  else assert.notEqual(child.status, 0, output);
  return output;
}

// CLI-ACTION-08 / SCN-CLI-ACTION-02: actual entrypoint help and error exits.
const rootHelp = run(["--help"]);
assert.doesNotMatch(rootHelp, /^\s+execute(?:\s|$)/m);
assert.match(rootHelp, /--interactive/);
assert.doesNotMatch(run(["execute", "--help"]), /--interactive|(?:^|\s)-i(?:\s|$)/);
const group = run(["execute", "action", "--help"]);
assert.match(group, /copilotAgent\/publishPolicy/);
assert.match(group, /teamsApp\/extendToM365/);
assert.doesNotMatch(group, /--interactive|(?:^|\s)-i(?:\s|$)/);
assert.match(group, /sign-in\s+prompts/);
const policy = run(["execute", "action", "copilotAgent/publishPolicy", "--help"]);
const policyCopy = policy.replace(/\s+/g, " ");
assert.match(
  policyCopy,
  /Publish a policy for an existing agent\. Currently supports Work IQ policies only\./
);
assert.match(policyCopy, /Publication does not verify runtime enforcement\./);
for (const flag of ["--title-id", "--policy-path", "--caller-id"]) {
  assert.ok(policy.includes(flag), policy);
}
assert.match(policy, /\[Required\]/);
assert.match(policyCopy, /Nonzero App ID GUID of the agent's runtime application/);
assert.match(policy, /real external writes/);
assert.doesNotMatch(policy, /--folder/);
assert.doesNotMatch(policy, /--interactive|(?:^|\s)-i(?:\s|$)/);
assert.match(policy, /current working directory/);
assert.doesNotMatch(policy, /project root/);
assert.match(policy, /sign-in\s+prompts/);
assert.match(run(["execute", "action", "teamsApp/extendToM365", "--help"]), /--app-package-path/);
assert.match(
  run(["execute", "action", "copilotAgent/publishPolicy"], false),
  /missing required option/i
);
assert.match(run(["execute", "action", "workiq/publishPolicy"], false), /not recognized/);
assert.match(run(["execute", "action", "arbitrary/module"], false), /not recognized/);
assert.match(
  run(
    [
      "execute",
      "action",
      "copilotAgent/publishPolicy",
      "--title-id",
      "U_title",
      "--policy-path",
      "unused-policy.json",
      "--caller-id",
      "   ",
    ],
    false
  ),
  /ActionArgsValidationError/
);
const folder = mkdtempSync(path.join(tmpdir(), "atk-cli-action-smoke-"));
try {
  const validArgs = [
    "execute",
    "action",
    "copilotAgent/publishPolicy",
    "--title-id",
    "U_title",
    "--policy-path",
    path.join(folder, "missing-policy.json"),
    "--caller-id",
    "11111111-2222-4333-8444-555555555555",
  ];
  for (const ci of ["false", "true"]) {
    const rejection = run([...validArgs, "--interactive", "true"], false, { CI_ENABLED: ci });
    assert.match(rejection, /ActionInteractiveUnsupported/);
    assert.match(rejection, /Action arguments/);
    assert.match(rejection, /configured token provider/);
  }
  for (const flags of [
    ["--policy-path=--debug"],
    ["--folder=--policy-path"],
    ["--policy-path=missing=policy.json"],
  ]) {
    assert.match(run([...validArgs, ...flags], false), /InvalidActionOptionValue/);
  }
  const failure = run(validArgs, false);
  assert.match(failure, /WorkIqPolicyReadFailed/);
  assert.match(failure, /"report":/);
  assert.match(failure, /"state": "not-started"/);
  const relativeArgs = [...validArgs];
  const cwd = path.join(folder, "cwd");
  mkdirSync(cwd);
  relativeArgs[relativeArgs.indexOf("--policy-path") + 1] = path.join("..", "missing-policy.json");
  assert.match(run(relativeArgs, false, {}, cwd), /WorkIqPolicyReadFailed/);
  assert.deepEqual(readdirSync(folder), ["cwd"]);
  assert.deepEqual(readdirSync(cwd), []);
} finally {
  rmSync(folder, { recursive: true, force: true });
}
process.stdout.write(
  "CLI-ACTION-08 / SCN-CLI-ACTION-02: built CLI help and error exits pass without publishing.\n"
);
