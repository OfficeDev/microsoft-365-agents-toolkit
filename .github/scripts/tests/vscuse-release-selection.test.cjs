const assert = require("node:assert/strict");
const { test } = require("node:test");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const downloader = path.resolve(__dirname, "../../actions/setup-vscuse-environment/download-vscuse-python.js");
const latestRoute = "GET /repos/{owner}/{repo}/releases/latest";
const tagRoute = "GET /repos/{owner}/{repo}/releases/tags/{tag}";
const listRoute = "GET /repos/{owner}/{repo}/releases";
const assetRoute = "GET /repos/{owner}/{repo}/releases/assets/{asset_id}";
const repository = { owner: "microsoft", repo: "vscuse-doc" };
const stableTag = "v0.2.67";
const rcTag = "v0.2.68-rc.1";
const wheel = { id: 42, name: "vscuse-0.2.67-py3-none-any.whl", browser_download_url: "https://example.test/wheel" };
const stable = {
  tag_name: stableTag,
  draft: false,
  prerelease: false,
  assets: [{ id: 41, name: "README.txt" }, wheel, { id: 43, name: "other.whl" }],
};
const rc = { ...stable, tag_name: rcTag, prerelease: true };

async function runMockDownloader() {
  const scenario = JSON.parse(fs.readFileSync(0, "utf8"));
  const trace = { requests: [], directories: [], writes: [] };
  process.on("exit", () => console.log(`TRACE:${JSON.stringify(trace)}`));
  const context = vm.createContext({
    Buffer,
    console,
    process: {
      argv: ["node", downloader, scenario.target, '"wheel output"'],
      env: { GH_APP_ID: "test-app", GH_APP_PRIVATE_KEY: "test\\nkey" },
      platform: process.platform,
      exit: (code) => process.exit(code),
    },
  });
  const modules = {
    fs: { default: { promises: {
      mkdir: async (directory, options) => trace.directories.push({ directory, options }),
      writeFile: async (destination, data) => trace.writes.push({ destination, data: data.toString() }),
    } } },
    path: { default: path },
    "node-fetch": { default: async (url, options) => {
      assert.equal(url, "https://api.github.com/app/installations");
      assert.equal(options.headers.Authorization, "Bearer test-app-token");
      return { json: async () => [{ id: 7, account: { login: "microsoft" } }] };
    } },
    "@octokit/auth-app": { createAppAuth: (options) => {
      assert.equal(options.appId, "test-app");
      assert.equal(options.privateKey, "test\nkey");
      return async ({ type, installationId }) => {
        if (type === "app") return { token: "test-app-token" };
        assert.equal(type, "installation");
        assert.equal(installationId, 7);
        return { token: "test-installation-token" };
      };
    } },
    "@octokit/core": { Octokit: class {
      constructor(options) {
        assert.equal(options.auth, "test-installation-token");
      }
      async request(route, options) {
        trace.requests.push({ route, options });
        assert.equal(options.owner, repository.owner);
        assert.equal(options.repo, repository.repo);
        if (route === listRoute) return { data: scenario.list };
        if (route === latestRoute || route === tagRoute) {
          if (scenario.error) throw new Error(scenario.error);
          return { data: scenario.release };
        }
        assert.equal(route, assetRoute);
        return { data: Buffer.from("mock wheel contents") };
      }
    } },
  };
  const entry = new vm.SourceTextModule(fs.readFileSync(downloader, "utf8"), { context });
  await entry.link((specifier) => {
    assert.ok(Object.hasOwn(modules, specifier), `Unexpected import: ${specifier}`);
    const exports = modules[specifier];
    return new vm.SyntheticModule(Object.keys(exports), function () {
      for (const [name, value] of Object.entries(exports)) this.setExport(name, value);
    }, { context });
  });
  await entry.evaluate();
}

function run(scenario = {}) {
  const result = spawnSync(process.execPath, ["--experimental-vm-modules", __filename, "--mock-downloader"], {
    input: JSON.stringify({ target: "latest", release: stable, list: [rc, stable], ...scenario }),
    encoding: "utf8",
    timeout: 10000,
  });
  if (result.error) throw result.error;
  const traceLine = result.stdout.split(/\r?\n/).find((line) => line.startsWith("TRACE:"));
  assert.ok(traceLine, result.stdout + result.stderr);
  return { ...result, trace: JSON.parse(traceLine.slice("TRACE:".length)) };
}

function assertLookup(result, target) {
  assert.deepEqual(result.trace.requests[0], {
    route: !target || target === "latest" ? latestRoute : tagRoute,
    options: !target || target === "latest" ? repository : { ...repository, tag: target },
  });
  assert.ok(result.trace.requests.every(({ route }) => route !== listRoute));
}

function assertDownloaded(result, target = "latest") {
  assert.equal(result.status, 0, result.stdout + result.stderr);
  assertLookup(result, target);
  assert.deepEqual(result.trace.requests.slice(1), [{
    route: assetRoute,
    options: { ...repository, asset_id: wheel.id, headers: { Accept: "application/octet-stream" } },
  }]);
  const directory = process.platform === "win32" ? path.resolve("wheel output") : "wheel output";
  assert.deepEqual(result.trace.directories, [{ directory, options: { recursive: true } }]);
  assert.deepEqual(result.trace.writes, [{
    destination: path.join(directory, wheel.name),
    data: "mock wheel contents",
  }]);
}

function assertFailed(result, target, message) {
  assert.equal(result.status, 1, result.stdout + result.stderr);
  assertLookup(result, target);
  assert.equal(result.trace.requests.length, 1, "Must not fall back or request an asset");
  assert.deepEqual(result.trace.directories, []);
  assert.deepEqual(result.trace.writes, []);
  if (message) assert.ok(result.stderr.includes(message), result.stderr);
}

if (process.argv[2] === "--mock-downloader") {
  runMockDownloader().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
} else {
  test("latest uses the stable endpoint regardless of release list order or first-page contents", () => {
    for (const list of [[rc, stable], [stable, rc], Array(30).fill(rc), []]) {
      const result = run({ list });
      assertDownloaded(result);
      assert.ok(result.stdout.includes(`Using latest release: ${stableTag}`));
    }
  });

  test("omitted and empty versions retain the stable latest default", () => {
    for (const target of [undefined, ""]) assertDownloaded(run({ target }), target);
  });

  for (const [name, release] of [
    ["prerelease", rc],
    ["draft", { ...stable, draft: true }],
    ["draft prerelease", { ...rc, draft: true }],
    ["missing prerelease flag", { ...stable, prerelease: undefined }],
    ["missing draft flag", { ...stable, draft: undefined }],
    ["non-boolean prerelease flag", { ...stable, prerelease: "false" }],
    ["non-boolean draft flag", { ...stable, draft: 0 }],
    ["missing tag", { ...stable, tag_name: undefined }],
    ["empty tag", { ...stable, tag_name: "" }],
    ["non-string tag", { ...stable, tag_name: 267 }],
    ["empty response", undefined],
    ["null response", null],
    ["array response", [stable]],
    ["missing assets", { ...stable, assets: undefined }],
    ["no wheel", { ...stable, assets: [{ name: "README.txt" }] }],
  ]) {
    test(`latest fails closed for ${name} without selecting another release`, () => {
      assertFailed(run({ release }), "latest");
    });
  }

  test("explicit stable and RC tags resolve exactly even when absent from the first page", () => {
    for (const release of [stable, rc]) {
      const result = run({ target: release.tag_name, release, list: Array(30).fill({ ...stable, tag_name: "v9.9.9" }) });
      assertDownloaded(result, release.tag_name);
      assert.ok(result.stdout.includes(`Using specified release: ${release.tag_name}`));
    }
  });

  for (const [name, release] of [
    ["a different RC", { ...rc, tag_name: "v0.2.68-rc.2" }],
    ["stable instead of RC", stable],
    ["a missing tag", { ...rc, tag_name: undefined }],
    ["an empty response", undefined],
    ["a null response", null],
    ["a list response", [rc]],
    ["missing assets", { ...rc, assets: undefined }],
  ]) {
    test(`explicit RC fails closed for ${name} without selecting another release`, () => {
      assertFailed(run({ target: rcTag, release, list: [rc, stable] }), rcTag);
    });
  }

  for (const target of ["latest", rcTag]) {
    for (const error of ["HTTP 404 Not Found", "HTTP 403 Forbidden", "HTTP 500 Server Error", "Network unavailable"]) {
      test(`${target} lookup failure (${error}) exits nonzero without fallback`, () => {
        assertFailed(run({ target, error, list: [rc, stable] }), target, error);
      });
    }
  }
}
