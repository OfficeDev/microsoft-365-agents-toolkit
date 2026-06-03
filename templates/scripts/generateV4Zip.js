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
 *   - build/v4/floor.json     — { "version": <templates package version> }
 *
 * The digest is NOT written here: it is computed from the bytes at load time
 * so `computeDigest` stays the single authority (resolve-template-source spec
 * decision #6). `loadBundledFloor` reads these two files.
 */

const AdmZip = require("adm-zip");
const { readdirSync, mkdirSync, writeFileSync } = require("node:fs");
const path = require("path");

const LANGUAGES = ["common", "js", "ts", "python"];
const BUILD_PATH = path.join(__dirname, "..", "build", "v4");
const version = require(path.join(__dirname, "..", "package.json")).version;

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
