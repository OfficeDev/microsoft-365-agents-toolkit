// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Generate the single full v4 template package + bundled-floor manifest.
 *
 * Unlike the v3 per-language zips (common/js/ts/python.zip), the v4 channel
 * ships ONE `templates.zip` (ADR-0006 / scaffolding.create.proposal.md). This
 * script bundles the existing VSC template content under `<lang>/<scenario>/`
 * paths into a single zip. The v4 `<templateId>` authoring layout is a later
 * concern (the consume operation); this script only produces the distribution
 * artifact the floor + `templates-v4@` channel ship.
 *
 * Outputs (both picked up by the `distribute` step → packages/fx-core/templates/v4/):
 *   - build/v4/templates.zip  — the full package
 *   - build/v4/floor.json     — { "version": <clean v4 publish version> }
 *
 * The floor version is the SAME clean, suffix-free version published to the v4
 * channel (computeV4PublishVersion / templates-config.json v4.localVersion), not
 * the raw minted package version: an odd-minor preview is minted as
 * `6.11.1-beta.<date>` but ships everywhere as `6.11.<date>`, so the bundled
 * floor and the online release report one identical version. Runs at build time
 * (lerna's `postversion` hook), so the minted version is already in package.json.
 *
 * The digest is NOT written here: it is computed from the bytes at load time
 * so `computeDigest` stays the single authority (resolve-template-source spec
 * decision #6). `loadBundledFloor` reads these two files.
 */

const AdmZip = require("adm-zip");
const { readdirSync, mkdirSync, writeFileSync } = require("node:fs");
const path = require("path");
const semver = require("semver");

// Mirror of packages/fx-core/src/v4/distribution/templateConfig.ts
// `computeV4PublishVersion` (canonical, unit-tested). Kept inline so the
// templates build needs no fx-core build output. Odd minor (prerelease) stamps
// the build date into the patch (6.11.<date>, read from the -beta.<date> preid);
// even minor (stable) uses major.minor.patch as-is. Keep in sync with canonical.
function computeV4PublishVersion(rawVersion) {
  const parsed = semver.parse(rawVersion);
  if (parsed === null) {
    throw new Error(`Cannot compute v4 publish version: "${rawVersion}" is not valid SemVer.`);
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

const LANGUAGES = ["common", "js", "ts", "python"];
const BUILD_PATH = path.join(__dirname, "..", "build", "v4");
const rawVersion = require(path.join(__dirname, "..", "package.json")).version;
const version = computeV4PublishVersion(rawVersion);

mkdirSync(BUILD_PATH, { recursive: true });

const zip = new AdmZip();
LANGUAGES.forEach((lang) => {
  const langPath = path.join(__dirname, "..", "vsc", lang);
  readdirSync(langPath).forEach((scenario) => {
    zip.addLocalFolder(path.join(langPath, scenario), path.posix.join(lang, scenario));
  });
});

console.log(`Generating v4 templates.zip (version ${version})`);
zip.writeZip(path.join(BUILD_PATH, "templates.zip"));

writeFileSync(path.join(BUILD_PATH, "floor.json"), JSON.stringify({ version }, null, 2) + "\n");
console.log(`Wrote v4 floor.json (version ${version})`);
