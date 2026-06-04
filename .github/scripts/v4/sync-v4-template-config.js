// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * CD wrapper (v4-isolated): write the `v4` distribution block into
 * templates-config.json from the freshly minted templates version and the
 * `goproduct` flag.
 *
 * This is the thin node glue around the unit-tested pure logic in
 * packages/fx-core/src/v4/distribution/templateConfig.ts. It does only IO:
 * read version + flag + previous config, call computeV4TemplateConfig, write
 * back. The `v4` block is fully isolated from v3's `version` / `useLocalTemplate`
 * (different semantics, v3 frozen) — see scaffolding.create.proposal.md §5.1.
 *
 * Usage (run AFTER `lerna version`, with the templates package version minted):
 *   PRODUCTION=<true|false> node .github/scripts/v4/sync-v4-template-config.js
 *
 * Requires fx-core to be built (build/v4/distribution/templateConfig.js).
 */

const path = require("path");
const fs = require("fs");

const repoRoot = path.join(__dirname, "../../..");
const templateVersion = require(path.join(repoRoot, "templates/package.json")).version;

const fxCorePath = path.join(repoRoot, "packages/fx-core");
const configFile = path.join(fxCorePath, "src/common/templates-config.json");
const { computeV4TemplateConfig } = require(path.join(
  fxCorePath,
  "build/v4/distribution/templateConfig.js"
));

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
