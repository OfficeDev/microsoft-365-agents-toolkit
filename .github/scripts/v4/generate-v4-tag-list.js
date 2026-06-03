// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * CD wrapper (v4-isolated): compute the digest of the v4 `templates.zip` and
 * append a `{ version, digest }` entry to the v4 channel's NDJSON tag-list.
 *
 * The v4 channel tag-list is one JSON object per line (NDJSON), unlike v3's
 * newline-separated git-tag text — because v4 carries a content digest that is
 * not derivable from git tags. `parseTagList` in
 * packages/fx-core/src/v4/distribution/templateSourcePort.ts reads exactly this
 * format. The digest is `sha256:<hex>` over the raw zip bytes, matching
 * `computeDigest` in templateSource.ts (the single digest authority).
 *
 * Re-running for an already-published version replaces that version's line
 * (idempotent), so a CD re-run does not duplicate entries.
 *
 * Usage (run AFTER the templates build, with build/v4/templates.zip present):
 *   node .github/scripts/v4/generate-v4-tag-list.js \
 *     --zip <path to templates.zip> \
 *     --version <semver> \
 *     --ndjson <path to read existing + write merged NDJSON>
 *
 * The --ndjson file may be absent (first release); it is created.
 */

const crypto = require("crypto");
const fs = require("fs");

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 2) {
    const key = argv[i];
    const value = argv[i + 1];
    if (!key || !key.startsWith("--") || value === undefined) {
      throw new Error(`Malformed argument near "${key}". Expected --flag value pairs.`);
    }
    args[key.slice(2)] = value;
  }
  return args;
}

const { zip, version, ndjson } = parseArgs(process.argv.slice(2));
if (!zip || !version || !ndjson) {
  throw new Error("Missing required argument. Need --zip, --version and --ndjson.");
}

const bytes = fs.readFileSync(zip);
const digest = "sha256:" + crypto.createHash("sha256").update(bytes).digest("hex");

// Read the existing NDJSON (if any) and index by version so a re-run replaces
// rather than duplicates. Missing file = first release.
const entries = new Map();
let existing = "";
try {
  existing = fs.readFileSync(ndjson, "utf8");
} catch (error) {
  if (error.code !== "ENOENT") {
    throw error;
  }
}
for (const line of existing.split("\n")) {
  const trimmed = line.replace(/\r$/, "").trim();
  if (trimmed === "") {
    continue;
  }
  const parsed = JSON.parse(trimmed);
  entries.set(parsed.version, parsed);
}

entries.set(version, { version, digest });

const merged = [...entries.values()].map((entry) => JSON.stringify(entry)).join("\n") + "\n";
fs.writeFileSync(ndjson, merged);

console.log(
  `================== v4 tag-list entry: ${JSON.stringify({ version, digest })} (${entries.size} total) ==================`
);
