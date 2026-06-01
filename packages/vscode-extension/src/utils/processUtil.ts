// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { exec } from "child_process";
import * as os from "os";
import kill from "tree-kill";

export const execModule = {
  exec: exec,
};

export const killModule = {
  killTree: kill,
};

class ProcessUtil {
  // kill process and its child processes
  async killProcess(pid: number, timeout = 5000, silent = true): Promise<void> {
    const tPromise = timeoutPromise(timeout);
    const killPromise = new Promise<void>((resolve, reject) => {
      killModule.killTree(pid, "SIGTERM", (err) => {
        if (err && !silent) {
          reject(err);
        } else {
          resolve();
        }
      });
    });
    await Promise.race([tPromise, killPromise]);
  }

  async getProcessIdsByPort(port: number): Promise<number[]> {
    const platform = os.platform();
    let command: string;
    if (platform === "win32") {
      command = `netstat -ano | findstr LISTENING | findstr :${port}`;
    } else if (platform === "darwin") {
      command = `lsof -i :${port} -sTCP:LISTEN -t`;
    } else {
      command = `lsof -i :${port} -sTCP:LISTEN -t 2>/dev/null || ss -tlnp sport = :${port}`;
    }

    return new Promise<number[]>((resolve) => {
      execModule.exec(command, { timeout: 5000 }, (error, stdout) => {
        if (error || !stdout.trim()) {
          resolve([]);
          return;
        }
        const pids = new Set<number>();
        if (platform === "win32") {
          const portPattern = new RegExp(`:(${port})\\s`, "g");
          for (const line of stdout.trim().split("\n")) {
            if (!portPattern.test(line)) continue;
            portPattern.lastIndex = 0;
            const parts = line.trim().split(/\s+/);
            const pid = parseInt(parts[parts.length - 1], 10);
            if (pid > 0) pids.add(pid);
          }
        } else {
          for (const line of stdout.trim().split("\n")) {
            const directPid = parseInt(line.trim(), 10);
            if (directPid > 0) {
              pids.add(directPid);
              continue;
            }
            // ss output: parse pid= from the last column
            const match = /pid=(\d+)/.exec(line);
            if (match) {
              const pid = parseInt(match[1], 10);
              if (pid > 0) pids.add(pid);
            }
          }
        }
        resolve(Array.from(pids));
      });
    });
  }
}

export function timeoutPromise(timeout: number, silent = true): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    setTimeout(() => {
      if (silent) resolve();
      else reject(new Error("Operation timeout"));
    }, timeout);
  });
}
export const processUtil = new ProcessUtil();
