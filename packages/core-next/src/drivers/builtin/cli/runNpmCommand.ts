// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as childProcess from "node:child_process";
import * as path from "node:path";
import * as os from "node:os";
import { ok, err } from "neverthrow";
import { z } from "zod";
import { createDriver } from "../../createDriver";
import { userError, systemError } from "../../../core/error";
import type { AtkError } from "../../../core/error";
import { parseOutputDirectives } from "../script/run";

const inputSchema = z.object({
  /** npm sub-command arguments, e.g. "install" or "run build --if-present" */
  args: z.string().min(1, "args must be a non-empty npm sub-command"),
  /** Working directory (relative to projectPath or absolute). Defaults to "./" */
  workingDirectory: z.string().optional(),
  /** Optional path to prepend to PATH when locating the npm binary */
  execPath: z.string().optional(),
});

type NpmRunConfig = z.infer<typeof inputSchema>;

export const runNpmCommandDriver = createDriver<NpmRunConfig>({
  id: "cli/runNpmCommand",
  name: "Run npm Command",
  inputSchema,
  execute: async (ctx, config) => {
    const projectPath = ctx.projectPath ?? process.cwd();

    const cwd = config.workingDirectory
      ? path.isAbsolute(config.workingDirectory)
        ? config.workingDirectory
        : path.resolve(projectPath, config.workingDirectory)
      : projectPath;

    const command = `npm ${config.args}`;
    const shell = defaultShell();

    ctx.logger.info(`[cli/runNpmCommand] Running: ${command}`);
    ctx.logger.debug(`[cli/runNpmCommand] cwd=${cwd}`);

    let stdout: string;
    let stderr: string;
    try {
      const env = config.execPath
        ? { ...process.env, PATH: `${config.execPath}${path.delimiter}${process.env.PATH ?? ""}` }
        : undefined;
      const result = await executeCommand(command, { cwd, shell, env });
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError) {
      return err(execError as AtkError);
    }

    if (stdout) {
      ctx.logger.debug(`[cli/runNpmCommand] stdout:\n${stdout}`);
    }
    if (stderr) {
      ctx.logger.warning(`[cli/runNpmCommand] stderr:\n${stderr}`);
    }

    const outputs = parseOutputDirectives(stdout);
    return ok({ outputs });
  },
});

function executeCommand(
  command: string,
  options: { cwd: string; shell: string; env?: NodeJS.ProcessEnv }
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    childProcess.exec(
      command,
      {
        cwd: options.cwd,
        shell: options.shell,
        env: options.env,
        encoding: "utf-8",
        maxBuffer: 10 * 1024 * 1024,
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

function convertExecError(
  command: string,
  error: childProcess.ExecException,
  stderr: string
): AtkError {
  if (error.killed) {
    return userError("NpmTimeoutError", `npm command timed out and was killed: ${command}`, {
      source: "cli/runNpmCommand",
      inner: error,
    });
  }
  const message = stderr?.trim() || error.message;
  return systemError(
    "NpmExecutionError",
    `npm command failed (exit code ${error.code ?? "unknown"}): ${message}`,
    { source: "cli/runNpmCommand", inner: error }
  );
}

function defaultShell(): string {
  if (process.env.SHELL) {
    return process.env.SHELL;
  }
  return os.platform() === "win32" ? "cmd.exe" : "/bin/sh";
}
