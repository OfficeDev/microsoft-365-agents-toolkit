const assert = require("node:assert/strict");
const { test } = require("node:test");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "../../..");
const scripts = path.join(root, ".github", "scripts");
const templates = path.join(root, "templates");
const targets = require(path.join(templates, "package.json")).templatesDependOnSDK;

test("SDK sync targets are existing active template directories", () => {
  for (const target of targets) {
    assert.ok(target.startsWith("vsc/"), `Not an active VS Code template: ${target}`);
    assert.ok(fs.statSync(path.join(templates, target)).isDirectory(), target);
  }
});

test("SDK sync completes with the current template list", () => {
  const filename = path.join(scripts, "sync-version.js");
  const writes = [];
  vm.runInNewContext(fs.readFileSync(filename, "utf8"), {
    __dirname: scripts,
    __filename: filename,
    require(name) {
      if (name === "fs") {
        return { ...fs, writeFileSync: (...args) => writes.push(args) };
      }
      return require(name);
    },
    console: { log() {} },
  }, { filename });
  assert.deepEqual(writes, []);
});

function runHook(mode, failCommand = "") {
  const result = spawnSync("bash", ["-c", `
record() {
  printf 'CALL %s\\n' "$1"
  if [ "$1" = "$FAIL_COMMAND" ]; then
    printf 'Injected failure: %s\\n' "$1" >&2
    return 42
  fi
}
node() { record "node \${2:-$(basename "$1")}"; }
git() { record "git $1"; }
export -f record node git
bash "$VERSION_SCRIPT" "$VERSION_MODE"
`], {
    cwd: root,
    env: {
      ...process.env,
      VERSION_SCRIPT: path.join(scripts, "pkg-version.sh").replaceAll("\\", "/"),
      VERSION_MODE: mode,
      FAIL_COMMAND: failCommand,
    },
    encoding: "utf8",
    timeout: 10000,
  });
  if (result.error) throw result.error;
  return {
    ...result,
    calls: result.stdout.split(/\r?\n/).filter(line => line.startsWith("CALL "))
      .map(line => line.slice(5)),
  };
}

for (const [mode, calls] of [
  ["template-sync", ["node sync-version.js", "git add"]],
  ["core-template", ["node syncVersion", "git add", "node updateUseLocalFlag", "git update-index"]],
]) {
  test(`${mode} preserves the successful versioning sequence`, () => {
    const result = runHook(mode);
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(result.calls, calls);
  });
}

for (const [mode, failCommand, calls] of [
  ["template-sync", "node sync-version.js", ["node sync-version.js"]],
  ["core-template", "node syncVersion", ["node syncVersion"]],
  ["core-template", "git add", ["node syncVersion", "git add"]],
  ["core-template", "node updateUseLocalFlag", ["node syncVersion", "git add", "node updateUseLocalFlag"]],
]) {
  test(`${mode} propagates ${failCommand} failure without running later commands`, () => {
    const result = runHook(mode, failCommand);
    assert.equal(result.status, 42, result.stdout + result.stderr);
    assert.ok(result.stderr.includes(`Injected failure: ${failCommand}`));
    assert.deepEqual(result.calls, calls);
  });
}
