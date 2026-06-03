// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Copies skill files from packages/vscode-extension/skills/microsoft-365-agents-toolkit
 * into this package's skill/ directory, run automatically before npm publish (prepack).
 */

"use strict";

const fs = require("fs");
const path = require("path");

const src = path.resolve(
  __dirname,
  "../../vscode-extension/skills/microsoft-365-agents-toolkit"
);
const dest = path.resolve(__dirname, "../skill");

if (!fs.existsSync(src)) {
  console.warn(`Warning: skill source not found at:\n  ${src}`);
  console.warn("Skipping — skill/ will not be updated.");
  process.exit(0);
}

fs.rmSync(dest, { recursive: true, force: true });
fs.cpSync(src, dest, { recursive: true });

console.log("✅ Skill files synced:");
console.log(`   ${src}`);
console.log(`→  ${dest}`);
