// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import kill from "tree-kill";

export const killModule = {
  killTree: kill,
};

class ProcessUtil {
  // kill process and its child processes
  async killProcess(pid: number, timeout = 5000, silent = true): Promise<void> {
    const tPromise = timeoutPromise(timeout, silent);
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
