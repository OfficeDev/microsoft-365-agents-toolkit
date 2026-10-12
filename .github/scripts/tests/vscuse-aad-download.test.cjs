const assert = require("node:assert/strict");
const { test } = require("node:test");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createHash } = require("node:crypto");

const source = fs.readFileSync(
  path.join(__dirname, "../../actions/setup-vscuse-environment/download-vscuse-python.js"),
  "utf8",
).replace(/^import .*;\r?$/gm, "")
  .replace("(async () => {", "globalThis.finished = (async () => {");
const branch = "refs/heads/qfai/vscuse-aad-prerelease-validation";
const wheel = Buffer.from("synthetic-wheel-for-offline-tests");
const digest = `sha256:${createHash("sha256").update(wheel).digest("hex")}`;
const provenance = {
  source_release_id: 410362563,
  source_repo: "1openwindow/vsc-use",
  source_tag: "v1.0.3rc1",
};
const marker = (value) => `<!-- vscuse-mirror:${JSON.stringify(value)} -->`;

async function download({ ref = branch, tag = "v1.0.3rc1", release = {}, asset = {}, data = wheel, sha256 = digest.slice(7) } = {}) {
  const requests = [];
  const writes = [];
  const errors = [];
  let exitCode = 0;
  const candidate = {
    tag_name: tag,
    prerelease: true,
    draft: false,
    target_commitish: "main",
    body: marker(provenance),
    assets: [{ id: 1, name: `vscuse-${tag.slice(1)}-py3-none-any.whl`, digest, ...asset }],
    ...release,
  };
  const context = vm.createContext({
    Buffer, path, createHash,
    process: {
      env: { GH_APP_ID: "fixture", GH_APP_PRIVATE_KEY: "fixture", GITHUB_REF: ref, VSCUSE_SHA256: sha256 },
      argv: ["node", "download.js", tag, "wheels"],
      platform: process.platform,
      exit: (code) => { exitCode = code; },
    },
    console: { log() {}, error: (...args) => errors.push(args.join(" ")) },
    fs: { promises: { mkdir: async () => {}, writeFile: async (...args) => writes.push(args) } },
    createAppAuth: () => async () => ({ token: "fixture-token" }),
    fetch: async () => ({ json: async () => [
      { id: 1, account: { login: "microsoft" } },
    ] }),
    Octokit: class {
      async request(route, args) {
        requests.push({ route, ...args });
        if (route.endsWith("/releases/tags/{tag}")) return { data: candidate };
        if (route.endsWith("/releases")) return { data: [candidate] };
        if (route.endsWith("/releases/assets/{asset_id}")) return { data };
        throw new Error("Unexpected request");
      }
    },
  });
  vm.runInContext(source, context);
  await context.finished;
  return { requests, writes, errors, exitCode };
}

test("isolated branch fetches the exact official prerelease and checks wheel digest", async () => {
  const result = await download();
  assert.equal(result.exitCode, 0, result.errors.join("\n"));
  assert.equal(result.requests[0].route, "GET /repos/{owner}/{repo}/releases/tags/{tag}");
  assert.ok(result.requests.every((request) => request.owner === "microsoft" && request.repo === "vscuse-doc"));
  assert.equal(result.requests[0].tag, "v1.0.3rc1");
  assert.equal(result.writes.length, 1);
  assert.deepEqual(result.writes[0][1], wheel);
});

test("latest and stable versions fail before download on the isolated branch", async () => {
  for (const tag of ["latest", "v1.0.3", "", "v1.0.3rc2", "v1.0.3rc1;unsafe", "v1.0.3rc1\n"]) {
    const result = await download({ tag });
    assert.equal(result.exitCode, 1, tag);
    assert.equal(result.requests.length, 0);
    assert.equal(result.writes.length, 0);
  }
});

test("draft, stable, malformed or mismatched releases fail without writing a wheel", async () => {
  for (const release of [
    { draft: true }, { draft: undefined }, { prerelease: false }, { prerelease: "true" },
    { tag_name: "v1.0.3rc2" },
  ]) {
    const result = await download({ release });
    assert.equal(result.exitCode, 1);
    assert.equal(result.writes.length, 0);
  }
});

test("mirror provenance is parsed as JSON regardless of property order or mirror commit", async () => {
  const result = await download({ release: {
    target_commitish: "a".repeat(40),
    body: `Original notes\n${marker({
      source_tag: provenance.source_tag,
      source_repo: provenance.source_repo,
      source_release_id: provenance.source_release_id,
    })}\n`,
  } });
  assert.equal(result.exitCode, 0, result.errors.join("\n"));
  assert.equal(result.writes.length, 1);
  assert.ok(result.requests.every((request) => request.owner === "microsoft" && request.repo === "vscuse-doc"));
});

test("missing, duplicate and malformed mirror markers fail before asset download", async () => {
  for (const body of [
    null, undefined, "", "Original notes only", 42,
    marker(provenance) + "\n" + marker(provenance),
    "<!-- vscuse-mirror:invalid-json -->",
    marker(null), marker([]), marker("text"), marker({}),
  ]) {
    const result = await download({ release: { body } });
    assert.equal(result.exitCode, 1);
    assert.equal(result.requests.length, 1);
    assert.equal(result.writes.length, 0);
  }
});

test("wrong source release, repo or tag fails even when wheel digest matches", async () => {
  for (const change of [
    { source_release_id: 410362564 },
    { source_release_id: "410362563" },
    { source_repo: "microsoft/vscuse-doc" },
    { source_tag: "v1.0.3rc2" },
  ]) {
    const result = await download({ release: { body: marker({ ...provenance, ...change }) } });
    assert.equal(result.exitCode, 1);
    assert.equal(result.requests.length, 1);
    assert.equal(result.writes.length, 0);
  }
});

test("the expected digest must be explicitly supplied by the release owner", async () => {
  for (const sha256 of ["", "a".repeat(63), "g".repeat(64), digest.slice(7) + "\n"]) {
    const result = await download({ sha256 });
    assert.equal(result.exitCode, 1);
    assert.equal(result.requests.length, 0);
    assert.equal(result.writes.length, 0);
  }
});

test("missing or wrong digest and wrong wheel versions fail closed", async () => {
  for (const asset of [
    { digest: null }, { digest: `sha256:${"0".repeat(64)}` },
    { name: "vscuse-1.0.3-py3-none-any.whl" },
    { name: "../vscuse-1.0.3rc1-py3-none-any.whl" },
  ]) {
    const result = await download({ asset });
    assert.equal(result.exitCode, 1);
    assert.equal(result.writes.length, 0);
  }
  const corrupt = await download({ data: Buffer.from("corrupted") });
  assert.equal(corrupt.exitCode, 1);
  assert.equal(corrupt.writes.length, 0);
});

test("ambiguous wheel assets fail closed", async () => {
  const result = await download({
    release: { assets: [
      { id: 1, name: "vscuse-1.0.3rc1-py3-none-any.whl", digest },
      { id: 2, name: "vscuse-1.0.3rc1-cp312-any.whl", digest },
    ] },
  });
  assert.equal(result.exitCode, 1);
  assert.equal(result.writes.length, 0);
});

test("other refs retain the existing daily-feed behavior", async () => {
  for (const ref of ["refs/heads/dev", "refs/heads/release/6.12", "refs/heads/unrelated"]) {
    const result = await download({ ref, tag: "latest", release: { tag_name: "v1.0.2" } });
    assert.equal(result.exitCode, 0, result.errors.join("\n"));
    assert.equal(result.requests[0].route, "GET /repos/{owner}/{repo}/releases");
    assert.ok(result.requests.every((request) => request.owner === "microsoft" && request.repo === "vscuse-doc"));
    assert.equal(result.writes.length, 1);
  }
});
