// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as childProcess from "node:child_process";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { ok, err } from "neverthrow";
import { z } from "zod";
import { createDriver } from "../../createDriver";
import { userError, systemError } from "../../../core/error";
import type { AtkError } from "../../../core/error";

const inputSchema = z.object({
  /** Shell command to execute */
  run: z.string().min(1, "run must be a non-empty command"),
  /** Working directory (relative to projectPath or absolute) */
  workingDirectory: z.string().optional(),
  /** Explicit shell binary path or name */
  shell: z.string().optional(),
  /** Timeout in milliseconds */
  timeout: z.number().positive().optional(),
  /** File path to redirect stdout/stderr output to */
  redirectTo: z.string().optional(),
});

type ScriptDriverConfig = z.infer<typeof inputSchema>;

/**
 * Driver: script
 *
 * Executes a shell command, captures stdout/stderr, and parses
 * `::set-output` / `::set-teamsfx-env` directives from stdout as outputs.
 *
 * Cross-platform: uses the system's default shell unless overridden.
 */
export const scriptDriver = createDriver<ScriptDriverConfig>({
  id: "script",
  name: "Run Script",
  inputSchema,
  execute: async (ctx, config) => {
    const projectPath = ctx.projectPath ?? process.cwd();

    const cwd = config.workingDirectory
      ? path.isAbsolute(config.workingDirectory)
        ? config.workingDirectory
        : path.resolve(projectPath, config.workingDirectory)
      : projectPath;

    const shell = config.shell || defaultShell();

    ctx.logger.info(`[script] Running: ${config.run}`);
    ctx.logger.debug(`[script] cwd=${cwd} shell=${shell}`);

    let stdout: string;
    let stderr: string;
    try {
      const result = await executeCommand(config.run, {
        cwd,
        shell,
        timeout: config.timeout,
      });
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError) {
      // executeCommand rejects with a typed AtkError
      return err(execError as AtkError);
    }

    // Log output
    if (stdout) {
      ctx.logger.debug(`[script] stdout:\n${stdout}`);
    }
    if (stderr) {
      ctx.logger.warning(`[script] stderr:\n${stderr}`);
    }

    // Redirect output to file if requested
    if (config.redirectTo) {
      const redirectPath = path.isAbsolute(config.redirectTo)
        ? config.redirectTo
        : path.resolve(projectPath, config.redirectTo);
      await fs.mkdir(path.dirname(redirectPath), { recursive: true });
      const combined = [stdout, stderr].filter(Boolean).join("\n");
      await fs.appendFile(redirectPath, combined + "\n", "utf-8");
    }

    // Parse ::set-output and ::set-teamsfx-env directives
    const outputs = parseOutputDirectives(stdout);

    return ok({ outputs });
  },
});

/**
 * Execute a shell command and return stdout/stderr as strings.
 */
function executeCommand(
  command: string,
  options: { cwd: string; shell: string; timeout?: number }
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    childProcess.exec(
      command,
      {
        cwd: options.cwd,
        shell: options.shell,
        timeout: options.timeout,
        encoding: "utf-8",
        maxBuffer: 10 * 1024 * 1024, // 10MB
      },
      (error, stdout, stderr) => {
        if (error) {
          reject(convertExecError(command, error, stderr));
          return;
        }
        resolve({ stdout: stdout ?? "", stderr: stderr ?? "" });
      }
    );
  });
}

/**
 * Convert a child_process exec error into a typed AtkError.
 */
function convertExecError(
  command: string,
  error: childProcess.ExecException,
  stderr: string
): AtkError {
  if (error.killed) {
    return userError("ScriptTimeoutError", `Script timed out and was killed: ${command}`, {
      source: "script",
      inner: error,
    });
  }
  const message = stderr?.trim() || error.message;
  return systemError(
    "ScriptExecutionError",
    `Script failed (exit code ${error.code ?? "unknown"}): ${message}`,
    { source: "script", inner: error }
  );
}

/**
 * Parse `::set-output` and `::set-teamsfx-env` directives from command output.
 *
 * Supported formats:
 *   ::set-output KEY=VALUE
 *   ::set-teamsfx-env KEY=VALUE
 *   ::set-output KEY="VALUE WITH SPACES"
 */
export function parseOutputDirectives(stdout: string): Record<string, string> {
  const outputs: Record<string, string> = {};
  const pattern = /^::set-(?:output|teamsfx-env)\s+(\w+)=(.*)$/gm;
  let match;
  while ((match = pattern.exec(stdout)) !== null) {
    const key = match[1];
    let value = match[2].trim();
    // Strip surrounding quotes
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    outputs[key] = value;
  }
  return outputs;
}

/**
 * Returns the default shell for the current platform.
 */
export function defaultShell(): string {
  if (process.env.SHELL) {
    return process.env.SHELL;
  }
  return os.platform() === "win32" ? "cmd.exe" : "/bin/sh";
}
