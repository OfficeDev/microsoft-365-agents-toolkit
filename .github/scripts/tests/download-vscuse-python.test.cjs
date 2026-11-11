const assert = require("node:assert/strict");
const { test } = require("node:test");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const filename = path.resolve(__dirname, "..", "..", "actions",
  "setup-vscuse-environment", "download-vscuse-python.js");
const source = fs.readFileSync(filename, "utf8")
  .replace(/^import [^\r\n]+ from "[^"]+";\r?$/gm, "");
const latestRoute = "GET /repos/{owner}/{repo}/releases/latest";
const tagRoute = "GET /repos/{owner}/{repo}/releases/tags/{tag}";
const assetRoute = "GET /repos/{owner}/{repo}/releases/assets/{asset_id}";
const repository = { owner: "microsoft", repo: "vscuse-doc" };
const workingDirectory = "E:\\mock-vscuse";
const storePath = path.win32.join(workingDirectory, "wheel output");

function release(tag, id, flags = {}) {
  return {
    tag_name: tag,
    draft: false,
    prerelease: false,
    ...flags,
    assets: [
      { id: id + 1, name: "vscuse.exe" },
      { id, name: `vscuse-${tag}-py3-none-any.whl` },
      { id: id + 2, name: "source.tar.gz" },
    ].map(asset => ({
      ...asset,
      browser_download_url: `https://example.invalid/assets/${asset.id}`,
    })),
  };
}

const prerelease = release("v3.0.0rc1", 30, { prerelease: true });
const draft = release("v3.0.0", 20, { draft: true });
const stable = release("v2.0.0", 10);
const olderStable = release("v1.0.0", 1);
const listedReleases = [prerelease, draft, stable];
const allReleases = [...listedReleases, olderStable];
const wheelContent = selected => Buffer.from(`wheel payload: ${selected.tag_name}`);

async function runDownloader(args = [], releaseError) {
  const result = {
    appAuthOptions: [], authCalls: [], fetchCalls: [], octokitOptions: [],
    requests: [], mkdirs: [], writes: [], errors: [], exitCodes: [],
  };
  const completion = vm.runInNewContext(source, {
    fs: {
      promises: {
        mkdir: async (directory, options) => {
          result.mkdirs.push({ directory, options: { ...options } });
        },
        writeFile: async (destination, content) => {
          result.writes.push({ destination, content });
        },
      },
    },
    path: {
      ...path.win32,
      resolve: (...parts) => path.win32.resolve(workingDirectory, ...parts),
    },
    fetch: async (url, options) => {
      result.fetchCalls.push({ url, ...options, headers: { ...options.headers } });
      return {
        json: async () => [
          { id: 1, account: { login: "another-owner" } },
          { id: 42, account: { login: "Microsoft" } },
        ],
      };
    },
    createAppAuth: options => {
      result.appAuthOptions.push({ ...options });
      return async request => {
        result.authCalls.push({ ...request });
        return { token: request.type === "app" ? "mock-app-auth" : "mock-installation-auth" };
      };
    },
    Octokit: class {
      constructor(options) {
        result.octokitOptions.push({ ...options });
      }

      async request(route, options) {
        const params = { ...options };
        if (options.headers) params.headers = { ...options.headers };
        result.requests.push({ route, params });
        // Let the old implementation download the wrong wheel, rather than fail on a mock route.
        if (route === "GET /repos/{owner}/{repo}/releases") {
          return { data: listedReleases };
        }
        if (route === latestRoute || route === tagRoute) {
          if (releaseError) throw releaseError;
          const selected = route === latestRoute ? stable
            : allReleases.find(item => item.tag_name === options.tag);
          if (!selected) throw Object.assign(new Error("Not Found"), { status: 404 });
          return { data: selected };
        }
        if (route === assetRoute) {
          const selected = allReleases.find(item => item.assets[1].id === options.asset_id);
          if (!selected) throw new Error(`Unexpected wheel asset: ${options.asset_id}`);
          return { data: Uint8Array.from(wheelContent(selected)).buffer };
        }
        throw new Error(`Unexpected request: ${route}`);
      }
    },
    process: {
      env: { GH_APP_ID: "12345", GH_APP_PRIVATE_KEY: "mock\\nprivate-key" },
      argv: ["node", filename, ...args],
      platform: "win32",
      exit: code => result.exitCodes.push(code),
    },
    console: {
      log() {},
      error: (...values) => result.errors.push(values),
    },
    Buffer,
  }, { filename, timeout: 1000 });
  assert.equal(typeof completion?.then, "function", "The full async IIFE must execute");
  await completion;
  return result;
}

function releaseRequest(tag) {
  return {
    route: tag === "latest" ? latestRoute : tagRoute,
    params: { ...repository, ...(tag === "latest" ? {} : { tag }) },
  };
}

function assertDownload(result, selected, directory, tag) {
  assert.deepEqual(result.exitCodes, [], result.errors.flat().join("\n"));
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.writes, [{
    destination: path.win32.join(directory, selected.assets[1].name),
    content: wheelContent(selected),
  }]);
  assert.deepEqual(result.mkdirs, [{ directory, options: { recursive: true } }]);
  assert.deepEqual(result.requests, [
    releaseRequest(tag),
    {
      route: assetRoute,
      params: {
        ...repository,
        asset_id: selected.assets[1].id,
        headers: { Accept: "application/octet-stream" },
      },
    },
  ]);
  assert.deepEqual(result.appAuthOptions, [{ appId: "12345", privateKey: "mock\nprivate-key" }]);
  assert.deepEqual(result.authCalls, [
    { type: "app" },
    { type: "installation", installationId: 42 },
  ]);
  assert.deepEqual(result.fetchCalls, [{
    url: "https://api.github.com/app/installations",
    headers: { Authorization: "Bearer mock-app-auth", Accept: "application/vnd.github+json" },
  }]);
  assert.deepEqual(result.octokitOptions, [{ auth: "mock-installation-auth" }]);
}

function assertFailure(result, tag, message) {
  assert.deepEqual(result.exitCodes, [1]);
  const errors = result.errors.flat().join("\n");
  assert.ok(errors.includes(message), errors);
  assert.deepEqual(result.mkdirs, []);
  assert.deepEqual(result.writes, []);
  assert.deepEqual(result.requests, [releaseRequest(tag)]);
}

test("AC-01/AC-05: omitted tag downloads latest stable, not the first prerelease or draft", async () => {
  const result = await runDownloader();
  assertDownload(result, stable, workingDirectory, "latest");
});

test("AC-01/AC-05: explicit latest downloads stable to the normalized quoted output path", async () => {
  const result = await runDownloader(["latest", `"${storePath}\\staging\\.."`]);
  assertDownload(result, stable, storePath, "latest");
});

test("AC-02/AC-05: an explicit prerelease resolves directly and downloads its wheel", async () => {
  const result = await runDownloader([prerelease.tag_name, `'${storePath}'`]);
  assertDownload(result, prerelease, storePath, prerelease.tag_name);
});

test("AC-02/AC-05: a stable tag outside the listed page resolves directly", async () => {
  const result = await runDownloader([olderStable.tag_name, storePath]);
  assertDownload(result, olderStable, storePath, olderStable.tag_name);
});

test("AC-03: missing tag reports a clear 404 failure without fallback or writes", async () => {
  const tag = "v0.0.0-missing";
  const result = await runDownloader([tag, storePath]);
  assertFailure(result, tag, `Release with tag '${tag}' not found`);
});

for (const tag of ["latest", prerelease.tag_name]) {
  for (const status of [403, 500]) {
    test(`AC-04: ${tag} propagates HTTP ${status} without fallback or writes`, async () => {
      const error = Object.assign(new Error(`Release lookup failed: HTTP ${status}`), { status });
      const result = await runDownloader([tag, storePath], error);
      assertFailure(result, tag, error.message);
    });
  }
}

test("AC-04: latest 404 propagates without falling back to the releases list", async () => {
  const error = Object.assign(new Error("Not Found"), { status: 404 });
  const result = await runDownloader(["latest", storePath], error);
  assertFailure(result, "latest", error.message);
});
