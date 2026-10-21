const assert = require("node:assert/strict");
const { test } = require("node:test");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const scriptPath = path.resolve(__dirname, "..", "apply-office-addin-prs.cjs");
const source = fs.readFileSync(scriptPath, "utf8");
const repo = "OfficeDev/microsoft-365-agents-toolkit";
const templateId = "office-addin-excel-customfunctions";
const botLogin = "teamstoolkitworkflowapp[bot]";

function pendingPr(overrides = {}) {
  return {
    state: "open",
    html_url: `https://github.com/${repo}/pull/16855`,
    user: { login: botLogin },
    base: { ref: "dev" },
    head: {
      ref: `auto/office-addin-drift/${templateId}-39213379532`,
      repo: { full_name: repo },
    },
    ...overrides,
  };
}

async function runScript({
  pages = [[]], status = 200, networkError, env = {},
  rows = [{ id: templateId, verdict: "update" }],
} = {}) {
  const requests = [];
  const commands = [];
  const logs = [];
  const errors = [];
  let copilotCalls = 0;
  const processStub = {
    env: {
      GH_TOKEN: "test-token",
      GITHUB_REPOSITORY: repo,
      PR_BASE_BRANCH: "dev",
      PR_BOT_LOGIN: botLogin,
      ...env,
    },
    exitCode: 0,
  };
  const fsStub = {
    existsSync: () => true,
    readFileSync: () => JSON.stringify({ rows }),
  };
  const childProcessStub = {
    execSync(command) {
      commands.push(command);
      return "";
    },
    execFileSync() {
      copilotCalls++;
      return "No changes needed.";
    },
  };
  await vm.runInNewContext(source, {
    __dirname: path.dirname(scriptPath),
    require(name) {
      if (name === "fs") return fsStub;
      if (name === "child_process") return childProcessStub;
      return require(name);
    },
    process: processStub,
    console: {
      log: (message) => logs.push(message),
      error: (message) => errors.push(message),
    },
    URL,
    AbortSignal,
    async fetch(input, options) {
      const url = new URL(input);
      if (!url.pathname.endsWith("/pulls")) return { ok: false };
      requests.push({ url, options });
      if (networkError) throw new Error(networkError);
      const page = Number(url.searchParams.get("page") || "1");
      return {
        ok: status === 200,
        status,
        json: async () => pages[page - 1] || [],
      };
    },
  }, { filename: scriptPath });
  return { requests, commands, logs, errors, copilotCalls, exitCode: processStub.exitCode };
}

test("matching open bot PR skips Copilot and all git operations, regardless of title", async () => {
  const result = await runScript({ pages: [[pendingPr({ title: "Edited by reviewer" })]] });
  assert.equal(result.exitCode, 0);
  assert.equal(result.copilotCalls, 0);
  assert.deepEqual(result.commands, []);
  assert.ok(result.logs.some((line) => line.includes("/pull/16855")));
  assert.equal(result.requests.length, 1);
  const { url, options } = result.requests[0];
  assert.equal(url.searchParams.get("state"), "open");
  assert.equal(url.searchParams.get("base"), "dev");
  assert.equal(url.searchParams.get("per_page"), "100");
  assert.equal(options.headers.Authorization, "Bearer test-token");
});

test("no open matching PR preserves the existing generation path", async () => {
  const result = await runScript();
  assert.equal(result.exitCode, 0);
  assert.equal(result.copilotCalls, 1);
  assert.ok(result.commands.includes("git checkout dev"));
  assert.ok(result.commands.includes("git reset --hard origin/dev"));
});

test("closed, merged, other-template, other-base, other-author and fork PRs do not block", async () => {
  for (const pr of [
    pendingPr({ state: "closed" }),
    pendingPr({ state: "closed", merged_at: "2026-10-19T00:00:00Z" }),
    pendingPr({ head: { ref: "auto/office-addin-drift/office-addin-outlook-taskpane-123",
      repo: { full_name: repo } } }),
    pendingPr({ base: { ref: "release" } }),
    pendingPr({ user: { login: "another-bot[bot]" } }),
    pendingPr({ head: { ref: `auto/office-addin-drift/${templateId}-123`,
      repo: { full_name: "someone/fork" } } }),
  ]) {
    const result = await runScript({ pages: [[pr]] });
    assert.equal(result.exitCode, 0);
    assert.equal(result.copilotCalls, 1);
  }
});

test("PR lookup follows pagination and finds an existing PR beyond the first page", async () => {
  const unrelated = Array.from({ length: 100 }, () => pendingPr({
    user: { login: "someone-else" },
  }));
  const result = await runScript({ pages: [unrelated, [pendingPr()]] });
  assert.equal(result.requests.length, 2);
  assert.equal(result.requests[1].url.searchParams.get("page"), "2");
  assert.equal(result.copilotCalls, 0);
  assert.deepEqual(result.commands, []);
});

test("skipping one template still allows another template to be processed", async () => {
  const result = await runScript({
    pages: [[pendingPr()]],
    rows: [
      { id: templateId, verdict: "update" },
      { id: "office-addin-outlook-taskpane", verdict: "update" },
    ],
  });
  assert.equal(result.exitCode, 0);
  assert.equal(result.requests.length, 2);
  assert.equal(result.copilotCalls, 1);
});

test("lookup uses the configured base branch and app bot identity", async () => {
  const result = await runScript({
    env: { PR_BASE_BRANCH: "release", PR_BOT_LOGIN: "configured-app[bot]" },
    pages: [[pendingPr({
      base: { ref: "release" },
      user: { login: "configured-app[bot]" },
    })]],
  });
  assert.equal(result.exitCode, 0);
  assert.equal(result.requests[0].url.searchParams.get("base"), "release");
  assert.equal(result.copilotCalls, 0);
  assert.deepEqual(result.commands, []);
});

test("lookup HTTP and network failures fail closed before Copilot or git", async () => {
  for (const options of [
    { status: 403 }, { status: 500 }, { networkError: "offline" },
    { pages: [{ message: "Unexpected response" }] },
  ]) {
    const result = await runScript(options);
    assert.equal(result.exitCode, 1);
    assert.equal(result.copilotCalls, 0);
    assert.deepEqual(result.commands, []);
    assert.ok(result.errors.length > 0);
  }
});

test("missing lookup credentials or identity fails closed", async () => {
  for (const key of ["GH_TOKEN", "GITHUB_REPOSITORY", "PR_BOT_LOGIN"]) {
    const result = await runScript({ env: { [key]: "" } });
    assert.equal(result.exitCode, 1);
    assert.equal(result.copilotCalls, 0);
    assert.deepEqual(result.commands, []);
    assert.ok(result.errors.some((line) => line.includes(key)));
  }
});

test("workflow serializes runs by target branch without cancelling active work", () => {
  const workflow = fs.readFileSync(path.resolve(
    __dirname, "..", "..", "workflows", "office-addin-manifest-drift-check.yml",
  ), "utf8").replaceAll("\r\n", "\n");
  assert.match(workflow, /\nconcurrency:\n  group: office-addin-manifest-drift-\$\{\{ github\.ref \}\}\n  cancel-in-progress: false\n/);
  assert.ok(workflow.includes("PR_BOT_LOGIN: ${{ steps.generate-token.outputs.app-slug }}[bot]"));
});
