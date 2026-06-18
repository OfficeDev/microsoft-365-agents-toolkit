// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Generate the single full v4 template package + bundled-floor manifest.
 *
 * Unlike the v3 per-language zips (common/js/ts/python.zip), the v4 channel
 * ships ONE `templates.zip` (ADR-0006 / scaffolding.create.proposal.md). This
 * script bundles two things into that single zip:
 *   1. the existing VSC template content under `<lang>/<scenario>/` paths
 *      (the v3 mirror, consumed by the transitional `{language, scenario}`
 *      locator), and
 *   2. the v4 authored packages under `v4/<kind>/<templateId>/`
 *      (descriptor/questions/pipeline JSON + a self-contained `content/`),
 *      consumed by the `{templateId}` locator once it ships.
 * v3 and v4 coexist; nothing here is stitched or synthesized — every byte is
 * authored (scaffolding.create.proposal.md §3 "authored, not generated").
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
const { readdirSync, mkdirSync, writeFileSync, existsSync } = require("node:fs");
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

// v4 authored packages (templates/v4/{create,modify}/...) ship self-contained:
// each package carries its own descriptor/questions/pipeline JSON AND a complete
// content/ copied from v3 during the v3↔v4 coexistence window. Build only zips
// the authored bytes verbatim under the `v4/` prefix — no stitching, no
// synthesis (scaffolding.create.proposal.md §3 "authored, not generated").
const v4SourcePath = path.join(__dirname, "..", "v4");
if (existsSync(v4SourcePath)) {
  zip.addLocalFolder(v4SourcePath, "v4");
}

console.log(`Generating v4 templates.zip (version ${version})`);
zip.writeZip(path.join(BUILD_PATH, "templates.zip"));

writeFileSync(path.join(BUILD_PATH, "floor.json"), JSON.stringify({ version }, null, 2) + "\n");
console.log(`Wrote v4 floor.json (version ${version})`);
