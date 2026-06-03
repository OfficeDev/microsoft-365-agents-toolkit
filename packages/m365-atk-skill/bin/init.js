#!/usr/bin/env node
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Microsoft 365 Agents Toolkit Skill Installer
 *
 * Installs the ATK skill into your project for your chosen AI coding tool:
 *   - Claude Code  → CLAUDE.md + .m365-atk-skill/
 *   - Cursor       → .cursor/rules/m365-atk.mdc + .cursor/m365-atk-skill/
 *   - GitHub Copilot → .github/copilot-instructions.md + .github/skills/microsoft-365-agents-toolkit/
 *   - OpenAI Codex → AGENTS.md + .m365-atk-skill/
 *
 * Usage:
 *   npx @microsoft/m365agentstoolkit-skill
 *   npx @microsoft/m365agentstoolkit-skill --tool claude-code
 *   npx @microsoft/m365agentstoolkit-skill --tool cursor
 *   npx @microsoft/m365agentstoolkit-skill --tool copilot
 *   npx @microsoft/m365agentstoolkit-skill --tool codex
 *   npx @microsoft/m365agentstoolkit-skill --tool all
 */

"use strict";

const fs = require("fs");
const path = require("path");
const readline = require("readline");

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const SKILL_SRC = path.join(__dirname, "..", "skill");

const MARKER_START = "<!-- m365-atk-skill-start -->";
const MARKER_END = "<!-- m365-atk-skill-end -->";

const TOOLS = {
  "claude-code": "Claude Code (claude.ai/code, Anthropic)",
  cursor: "Cursor (cursor.com)",
  copilot: "GitHub Copilot CLI / Coding Agent",
  codex: "OpenAI Codex / ChatGPT Coding Agent",
  all: "All of the above",
};

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function requireSkillSrc() {
  if (!fs.existsSync(SKILL_SRC)) {
    console.error(
      "❌  Skill files not found.\n" +
        "    If running from source, run: node scripts/copy-skill.js\n" +
        "    Expected location: " +
        SKILL_SRC
    );
    process.exit(1);
  }
}

function copySkillFiles(targetDir) {
  requireSkillSrc();
  fs.mkdirSync(targetDir, { recursive: true });
  fs.cpSync(SKILL_SRC, targetDir, { recursive: true });
}

/**
 * Append content wrapped in idempotency markers.
 * If the markers already exist the block is replaced (not duplicated).
 */
function upsertMarkedBlock(filePath, content) {
  const block = `${MARKER_START}\n${content}\n${MARKER_END}`;

  if (fs.existsSync(filePath)) {
    const existing = fs.readFileSync(filePath, "utf8");
    const re = new RegExp(
      `${escapeRe(MARKER_START)}[\\s\\S]*?${escapeRe(MARKER_END)}`,
      "g"
    );
    if (re.test(existing)) {
      fs.writeFileSync(filePath, existing.replace(re, block));
      console.log(`  ✅ Updated ATK skill block in ${filePath}`);
      return;
    }
    fs.appendFileSync(filePath, `\n\n${block}\n`);
    console.log(`  ✅ Appended ATK skill block to ${filePath}`);
  } else {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    fs.writeFileSync(filePath, `${block}\n`);
    console.log(`  ✅ Created ${filePath}`);
  }
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Read SKILL.md and strip the YAML frontmatter (--- … ---) if present. */
function readSkillBody() {
  const raw = fs.readFileSync(path.join(SKILL_SRC, "SKILL.md"), "utf8");
  return raw.replace(/^---[\s\S]*?---\n+/, "");
}

// ---------------------------------------------------------------------------
// Per-tool setup functions
// ---------------------------------------------------------------------------

function setupClaudeCode(cwd) {
  const skillDir = path.join(cwd, ".m365-atk-skill");
  copySkillFiles(skillDir);
  console.log(`  ✅ Skill files copied to .m365-atk-skill/`);

  upsertMarkedBlock(
    path.join(cwd, "CLAUDE.md"),
    "# Microsoft 365 Agents Toolkit Skill\n\n@./.m365-atk-skill/SKILL.md"
  );
}

function setupCursor(cwd) {
  const skillDir = path.join(cwd, ".cursor", "m365-atk-skill");
  copySkillFiles(skillDir);
  console.log(`  ✅ Skill files copied to .cursor/m365-atk-skill/`);

  const mdcPath = path.join(cwd, ".cursor", "rules", "m365-atk.mdc");
  if (fs.existsSync(mdcPath)) {
    console.log(`  ℹ️  .cursor/rules/m365-atk.mdc already exists — skipping`);
    return;
  }

  const body = readSkillBody();
  const mdc =
    `---\n` +
    `description: Microsoft 365 Agents Toolkit — build Teams apps and M365 agents\n` +
    `globs: ["m365agents*.yml", "teamsapp*.yml", "**/*.ts", "**/*.js", "**/*.py"]\n` +
    `alwaysApply: false\n` +
    `---\n\n` +
    body;

  fs.mkdirSync(path.dirname(mdcPath), { recursive: true });
  fs.writeFileSync(mdcPath, mdc);
  console.log(`  ✅ Created .cursor/rules/m365-atk.mdc`);
}

function setupCopilot(cwd) {
  const skillDir = path.join(
    cwd,
    ".github",
    "skills",
    "microsoft-365-agents-toolkit"
  );
  copySkillFiles(skillDir);
  console.log(
    `  ✅ Skill files copied to .github/skills/microsoft-365-agents-toolkit/`
  );

  upsertMarkedBlock(
    path.join(cwd, ".github", "copilot-instructions.md"),
    "# Microsoft 365 Agents Toolkit\n\n" +
      "Use the skill at `.github/skills/microsoft-365-agents-toolkit/SKILL.md` " +
      "when working on Teams apps or Microsoft 365 agent projects."
  );
}

function setupCodex(cwd) {
  const skillDir = path.join(cwd, ".m365-atk-skill");
  copySkillFiles(skillDir);
  console.log(`  ✅ Skill files copied to .m365-atk-skill/`);

  const body = readSkillBody();
  // readSkillBody() already starts with the top-level heading; use it directly
  upsertMarkedBlock(path.join(cwd, "AGENTS.md"), body);
}

function setupAll(cwd) {
  console.log("\n🔧 Setting up for Claude Code...");
  setupClaudeCode(cwd);
  console.log("\n🔧 Setting up for Cursor...");
  setupCursor(cwd);
  console.log("\n🔧 Setting up for GitHub Copilot...");
  setupCopilot(cwd);
  console.log("\n🔧 Setting up for OpenAI Codex...");
  setupCodex(cwd);
}

const TOOL_FNS = {
  "claude-code": setupClaudeCode,
  cursor: setupCursor,
  copilot: setupCopilot,
  codex: setupCodex,
  all: setupAll,
};

// ---------------------------------------------------------------------------
// Argument parsing
// ---------------------------------------------------------------------------

function parseArgs() {
  const args = process.argv.slice(2);
  let tool = null;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--tool" && args[i + 1]) {
      tool = args[++i];
    } else if (args[i].startsWith("--tool=")) {
      tool = args[i].split("=")[1];
    }
  }

  return { tool };
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const cwd = process.cwd();
  const { tool: toolArg } = parseArgs();

  // Non-interactive guard
  if (!toolArg && !process.stdin.isTTY) {
    console.error(
      "❌  stdin is not a TTY and --tool was not provided.\n" +
        "    Usage: npx @microsoft/m365agentstoolkit-skill --tool <tool>\n" +
        "    Tools: " +
        Object.keys(TOOLS).join(", ")
    );
    process.exit(1);
  }

  let toolKey = toolArg;

  if (!toolKey) {
    // Interactive prompt
    console.log("🚀  Microsoft 365 Agents Toolkit Skill Installer\n");
    console.log(
      "Which AI coding tool do you want to install the skill for?\n"
    );
    const toolKeys = Object.keys(TOOLS);
    toolKeys.forEach((key, i) => {
      console.log(`  ${i + 1}. ${TOOLS[key]}`);
    });

    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });

    toolKey = await new Promise((resolve) => {
      rl.question("\nEnter number or name: ", (answer) => {
        rl.close();
        const num = parseInt(answer.trim(), 10);
        if (!isNaN(num) && num >= 1 && num <= toolKeys.length) {
          resolve(toolKeys[num - 1]);
        } else {
          resolve(answer.trim().toLowerCase());
        }
      });
    });
  }

  if (!TOOL_FNS[toolKey]) {
    console.error(
      `❌  Unknown tool: "${toolKey}"\n` +
        "    Valid options: " +
        Object.keys(TOOLS).join(", ")
    );
    process.exit(1);
  }

  console.log(
    `\n📦  Installing for: ${TOOLS[toolKey]}\n    Target project: ${cwd}\n`
  );

  TOOL_FNS[toolKey](cwd);

  console.log("\n✅  Done! The Microsoft 365 Agents Toolkit skill is installed.");
  console.log(
    "    See https://github.com/OfficeDev/microsoft-365-agents-toolkit for documentation."
  );
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
