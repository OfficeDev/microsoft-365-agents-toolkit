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
  /** dotnet sub-command arguments, e.g. "publish --configuration Release" */
  args: z.string().min(1, "args must be a non-empty dotnet sub-command"),
  /** Working directory (relative to projectPath or absolute). Defaults to "./" */
  workingDirectory: z.string().optional(),
  /** Optional path to prepend to PATH when locating the dotnet binary */
  execPath: z.string().optional(),
});

type DotnetRunConfig = z.infer<typeof inputSchema>;

export const runDotnetCommandDriver = createDriver<DotnetRunConfig>({
  id: "cli/runDotnetCommand",
  name: "Run dotnet Command",
  inputSchema,
  execute: async (ctx, config) => {
    const projectPath = ctx.projectPath ?? process.cwd();

    const cwd = config.workingDirectory
      ? path.isAbsolute(config.workingDirectory)
        ? config.workingDirectory
        : path.resolve(projectPath, config.workingDirectory)
      : projectPath;

    const command = `dotnet ${config.args}`;
    const shell = defaultShell();

    ctx.logger.info(`[cli/runDotnetCommand] Running: ${command}`);
    ctx.logger.debug(`[cli/runDotnetCommand] cwd=${cwd}`);

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
      ctx.logger.debug(`[cli/runDotnetCommand] stdout:\n${stdout}`);
    }
    if (stderr) {
      ctx.logger.warning(`[cli/runDotnetCommand] stderr:\n${stderr}`);
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
    return userError("DotnetTimeoutError", `dotnet command timed out and was killed: ${command}`, {
      source: "cli/runDotnetCommand",
      inner: error,
    });
  }
  const message = stderr?.trim() || error.message;
  return systemError(
    "DotnetExecutionError",
    `dotnet command failed (exit code ${error.code ?? "unknown"}): ${message}`,
    { source: "cli/runDotnetCommand", inner: error }
  );
}

function defaultShell(): string {
  if (process.env.SHELL) {
    return process.env.SHELL;
  }
  return os.platform() === "win32" ? "cmd.exe" : "/bin/sh";
}
