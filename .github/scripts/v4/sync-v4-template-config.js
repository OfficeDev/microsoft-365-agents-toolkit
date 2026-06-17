// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * CD wrapper (v4-isolated): write the `v4` distribution block into
 * templates-config.json from the freshly minted templates version and the
 * `goproduct` flag.
 *
 * This does only IO: read version + flag + previous config, compute the `v4`
 * block, write back. The `v4` block is fully isolated from v3's `version` /
 * `useLocalTemplate` (different semantics, v3 frozen) — see
 * scaffolding.create.proposal.md §5.1.
 *
 * The compute logic below is intentionally self-contained plain JS and is NOT
 * required from fx-core's build output. A template-only release trims the pnpm
 * workspace to just `templates` (see .github/scripts/lernaDeps.json), so fx-core
 * is never built in that lane — requiring its compiled output would crash this
 * step. The canonical, unit-tested definition lives in
 * packages/fx-core/src/v4/distribution/templateConfig.ts; keep these in sync.
 *
 * `semver` is resolved from the `templates` package (always present when this
 * step runs — it only fires on a `templates@` change) because pnpm's per-package
 * node_modules (shared-workspace-lockfile=false) means a bare `require("semver")`
 * from this script's location does not resolve.
 *
 * Usage (run AFTER `lerna version`, with the templates package version minted):
 *   PRODUCTION=<true|false> node .github/scripts/v4/sync-v4-template-config.js
 */

const path = require("path");
const fs = require("fs");
const { createRequire } = require("module");

const repoRoot = path.join(__dirname, "../../..");
const templatesPackageJson = path.join(repoRoot, "templates/package.json");
const templateVersion = require(templatesPackageJson).version;
const semver = createRequire(templatesPackageJson)("semver");

const fxCorePath = path.join(repoRoot, "packages/fx-core");
const configFile = path.join(fxCorePath, "src/common/templates-config.json");

// Mirror of packages/fx-core/src/v4/distribution/templateConfig.ts (canonical,
// unit-tested). Kept inline so this CD step stays build-independent.
// `bundled` is just !goproduct. A shipping prerelease (odd-minor preview) is
// published online under a clean, suffix-free version (computeV4PublishVersion),
// so a suffix never forces bundling. Keep in sync with the canonical module.
function computeBundled(goproduct) {
  return !goproduct;
}

// Clean, suffix-free version published to the v4 channel: odd minor (prerelease)
// stamps the build date into the patch (6.11.<date>, read from the -beta.<date>
// preid); even minor (stable) uses major.minor.patch as-is.
function computeV4PublishVersion(version) {
  const parsed = semver.parse(version);
  if (parsed === null) {
    throw new Error(`Cannot compute v4 publish version: "${version}" is not valid SemVer.`);
  }
  if (parsed.minor % 2 === 1) {
    const dateStamp = parsed.prerelease.find(
      (segment) => typeof segment === "number" && segment >= 1000000000
    );
    if (dateStamp !== undefined) {
      return `${parsed.major}.${parsed.minor}.${dateStamp}`;
    }
  }
  return `${parsed.major}.${parsed.minor}.${parsed.patch}`;
}

function computeRange(version, previousRange) {
  const parsed = semver.parse(version);
  if (parsed === null) {
    throw new Error(`Cannot compute v4 template range: "${version}" is not valid SemVer.`);
  }
  if (semver.intersects(previousRange, `${parsed.major}.${parsed.minor}.0`)) {
    return previousRange;
  }
  return `~${parsed.major}.${parsed.minor}`;
}

function computeV4TemplateConfig(input) {
  return {
    range: computeRange(input.version, input.previousRange),
    bundled: computeBundled(input.goproduct),
    localVersion: computeV4PublishVersion(input.version),
  };
}

const goproduct = process.env.PRODUCTION === "true";

const config = JSON.parse(fs.readFileSync(configFile, "utf8"));
const previousRange = (config.v4 && config.v4.range) || config.version;

config.v4 = computeV4TemplateConfig({
  version: templateVersion,
  goproduct,
  previousRange,
});

// Match the existing 4-space, trailing-newline, LF convention of the file.
fs.writeFileSync(configFile, JSON.stringify(config, null, 4) + "\n");

console.log(
  `================== v4 template config: ${JSON.stringify(config.v4)} (goproduct=${goproduct}) ==================`
);
