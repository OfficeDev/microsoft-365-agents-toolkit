/* eslint-disable no-console */
"use strict";

const { spawnSync } = require("child_process");

function run(command, args) {
  const result = spawnSync(command, args, { encoding: "utf8" });

  if (result.error) {
    throw result.error;
  }

  if (typeof result.status === "number" && result.status !== 0) {
    throw new Error(result.stderr || `Command failed: ${command}`);
  }

  return (result.stdout || "").trim();
}

function getChangedSourceFiles(baseRef) {
  const output = run("git", ["diff", "--name-only", `${baseRef}...HEAD`]);
  return output
    .split(/\r?\n/)
    .map((file) => file.trim())
    .filter((file) => file.length > 0)
    .filter((file) => file.startsWith("src/"))
    .filter((file) => file.endsWith(".ts") || file.endsWith(".js"));
}

function main() {
  const baseRef = process.env.COVERAGE_BASE_REF || "origin/dev";

  let changedFiles = [];
  try {
    changedFiles = getChangedSourceFiles(baseRef);
  } catch (error) {
    console.error(`[coverage:changed] Failed to compute diff against ${baseRef}.`);
    console.error(String(error));
    process.exit(1);
  }

  if (changedFiles.length === 0) {
    console.log(`[coverage:changed] No changed source files found against ${baseRef}.`);
    process.exit(0);
  }

  console.log(`[coverage:changed] Base ref: ${baseRef}`);
  console.log("[coverage:changed] Files under coverage:");
  changedFiles.forEach((file) => console.log(` - ${file}`));

  const pnpmCmd = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
  const args = [
    "exec",
    "vitest",
    "run",
    "--coverage",
    "--changed",
    "--config",
    "vitest.config.ts",
    "--coverage.thresholds.lines",
    "0",
  ];

  changedFiles.forEach((file) => {
    args.push("--coverage.include", file);
  });

  const result = spawnSync(pnpmCmd, args, { stdio: "inherit" });
  if (typeof result.status === "number") {
    process.exit(result.status);
  }
  process.exit(1);
}

main();
