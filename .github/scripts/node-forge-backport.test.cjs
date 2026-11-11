// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const { createHash } = require("node:crypto");
const { readFileSync, realpathSync } = require("node:fs");
const { createRequire } = require("node:module");
const { dirname, join, resolve } = require("node:path");
const { test } = require("node:test");

const root = resolve(__dirname, "..", "..");
const coreRequire = createRequire(
  join(root, "packages", "fx-core", "package.json"),
);
const { parse } = coreRequire("yaml");
const identity = "node-forge@1.4.0";
const patchPath = "patches/node-forge@1.4.0.patch";
const patchSha256 =
  "1c99645b3297fbb3b77422e975aa2d397095bd1f5ce20b8fc6dc6da70141b50a";
const lockHash = "eoguvlerboigznr6skdgpf4qva";
const patchedVersion = `1.4.0(patch_hash=${lockHash})`;
const fileHashes = {
  "lib/rsa.js":
    "acc22e5d36e27832c34e02dd3933aad7977d45b047eead5016520735efedc9c5",
  "lib/pem.js":
    "38f22a503dc8285e31edfb67f0614a30c8af4abd0032e7738d45a08c31361401",
  "lib/pbe.js":
    "674750301f7c7d44b7b52c80bde90399aef8ff6021e6d6ccd1553ce3f5768c81",
  "lib/pkcs12.js":
    "4e471bf3e717d940a07b34d5716d988cfa01b143c38c3fc1f8b597cb83d29742",
  "lib/tls.js":
    "3b3f28f210e1911e03b43177576f689f38e4cc78207171fb245c9a1a60d1517b",
  "lib/x509.js":
    "cf3a94b7966de3445364ff20c08ec7810ea22a5222a54873f34df39e9c3db77f",
};
const sha256 = (file) =>
  createHash("sha256").update(readFileSync(file)).digest("hex");
const json = (file) => JSON.parse(readFileSync(file, "utf8"));
const normalize = (value) =>
  value
    .replace(/^\//, "")
    .replace(/ms-feed-\d+\.pkgs\.visualstudio\.com\//g, "");
const installedCopies = new Set();
for (const consumer of ["fx-core", "cli", "vscode-extension", "server"]) {
  const consumerRequire = createRequire(
    join(root, "packages", consumer, "package.json"),
  );
  const loader =
    consumer === "fx-core"
      ? consumerRequire
      : createRequire(
          consumerRequire.resolve("@microsoft/teamsfx-core/package.json"),
        );
  installedCopies.add(realpathSync(loader.resolve("node-forge/package.json")));
}

test("node-forge backport is content-bound in both manifests and every locked reference", () => {
  const manifest = json(join(root, "package.json"));
  const workspace = parse(
    readFileSync(join(root, "pnpm-workspace.yaml"), "utf8"),
  );
  const lock = parse(readFileSync(join(root, "pnpm-lock.yaml"), "utf8"));
  assert.equal(manifest.pnpm.patchedDependencies[identity], patchPath);
  assert.equal(workspace.patchedDependencies[identity], patchPath);
  assert.equal(
    sha256(join(root, patchPath)),
    patchSha256,
    "Patch content changed",
  );
  assert.deepEqual(lock.patchedDependencies[identity], {
    hash: lockHash,
    path: patchPath,
  });
  const nodes = Object.keys(lock.packages).filter((key) =>
    normalize(key).startsWith("node-forge@"),
  );
  assert.deepEqual(nodes.map(normalize), [`node-forge@${patchedVersion}`]);
  assert.equal(lock.packages[nodes[0]].patched, true);
  let references = 0;
  function verifyReferences(value) {
    if (!value || typeof value !== "object") return;
    for (const [name, entry] of Object.entries(value)) {
      if (name === "node-forge") {
        const reference = typeof entry === "string" ? entry : entry.version;
        assert.equal(
          reference,
          patchedVersion,
          "Unpatched node-forge lockfile reference",
        );
        references++;
      } else {
        verifyReferences(entry);
      }
    }
  }
  verifyReferences(lock.importers);
  verifyReferences(lock.packages);
  assert.ok(references > 0, "Missing node-forge dependency references");
});

for (const manifestPath of installedCopies) {
  const installed = dirname(manifestPath);
  test("core and shipping surfaces resolve the verified backported runtime", () => {
    assert.equal(json(manifestPath).name, "node-forge");
    assert.equal(json(manifestPath).version, "1.4.0");
    for (const [file, hash] of Object.entries(fileHashes)) {
      assert.equal(
        sha256(join(installed, file)),
        hash,
        `Installed runtime content changed: ${file}`,
      );
    }
  });

  const regressions = {
    "GHSA-86w9-cpqp-85rv": ["rsa"],
    "GHSA-7wch-p9px-956x": ["pem"],
    "GHSA-jc66-7q78-782g": ["pkcs12Kdf", "pkcs12Mac"],
    "GHSA-hqq3-7jhf-c5jx": ["tlsStack"],
    "GHSA-8hrf-8g6f-7xwp": ["tlsFragment"],
    "GHSA-5x38-2vff-x55p": ["x509"],
  };
  for (const [advisory, scenarios] of Object.entries(regressions)) {
    for (const scenario of scenarios) {
      test(`${advisory}: ${scenario}`, () => {
        // Isolate synchronous ReDoS/stack failures so a missing patch cannot hang CI.
        const result = spawnSync(
          process.execPath,
          [
            join(__dirname, "node-forge-regression.cjs"),
            join(installed, "lib", "index.js"),
            scenario,
          ],
          { encoding: "utf8", timeout: 15000, windowsHide: true },
        );
        assert.ifError(result.error);
        assert.equal(result.status, 0, result.stderr);
        assert.deepEqual(JSON.parse(result.stdout), { completed: scenario });
      });
    }
  }
}
