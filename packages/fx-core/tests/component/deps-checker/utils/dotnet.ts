// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as child_process from "child_process";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import * as tmp from "tmp";

import { ConfigFolderName } from "@microsoft/teamsfx-api";
import {
  DotnetChecker,
  DotnetVersion,
} from "../../../../src/component/deps-checker/internal/dotnetChecker";
import { cpUtils } from "../../../../src/component/deps-checker/util/cpUtils";
import { isArm64, isMacOS, isWindows } from "../../../../src/component/deps-checker/util/system";
import { logger } from "../adapters/testLogger";
import { createTmpDir } from "./common";

tmp.setGracefulCleanup();

export const dotnetConfigPath = path.join(os.homedir(), "." + ConfigFolderName, "dotnet.json");
export const dotnetPrivateInstallPath = path.join(
  os.homedir(),
  "." + ConfigFolderName,
  "bin",
  "dotnet"
);
export const dotnetCommand = "dotnet";
export const dotnetOldVersion = DotnetVersion.v21;
export const dotnetInstallVersion = isMacOS() && isArm64() ? DotnetVersion.v60 : DotnetVersion.v31;
export const dotnetSupportedVersions = [DotnetVersion.v31, DotnetVersion.v50];

export const testCsprojFileName = "extensions.csproj";
export const testOutputDirName = "bin";

export async function getDotnetExecPathFromConfig(
  dotnetConfigPath: string
): Promise<string | null> {
  try {
    const config = await fs.readJson(dotnetConfigPath, { encoding: "utf-8" });
    if (typeof config.dotnetExecutablePath === "string") {
      return config.dotnetExecutablePath;
    }
  } catch (error) {
    console.debug(`Failed to getDotnetConfig, error = '${error}'`);
  }
  return null;
}

export async function hasDotnetVersion(
  dotnetExecPath: string,
  versionString: string
): Promise<boolean> {
  return await hasAnyDotnetVersions(dotnetExecPath, [versionString]);
}

export async function hasAnyDotnetVersions(
  dotnetExecPath: string,
  versionStrings: string[]
): Promise<boolean> {
  try {
    const output = await cpUtils.executeCommand(
      undefined,
      logger,
      { shell: false },
      dotnetExecPath,
      "--list-sdks"
    );
    return output.split(/\r?\n/).some((line: string) => {
      return versionStrings.some((versionString) => line.startsWith(versionString));
    });
  } catch (error) {
    console.debug(`Failed to run "${dotnetExecPath} --list-sdks", error = '${error}'`);
    return false;
  }
}

export async function cleanup() {
  // fs-extra.remove() does nothing if the file does not exist.
  await fs.remove(dotnetConfigPath);
  try {
    if (isWindows()) {
      // Use tasklist instead of WMIC (removed in newer Windows versions)
      const output = child_process.execSync(
        'tasklist /FI "IMAGENAME eq dotnet.exe" /FO CSV /NH',
        { encoding: "utf-8" }
      );
      for (const line of output.split(/\r?\n/)) {
        const match = line.match(/^"[^"]*","(\d+)"/);
        if (match) {
          try {
            process.kill(Number(match[1]), "SIGKILL");
          } catch {
            // process may have already exited
          }
        }
      }
    } else {
      const output = child_process.execSync("pgrep -x dotnet || true", { encoding: "utf-8" });
      for (const line of output.trim().split(/\r?\n/)) {
        const pid = Number(line.trim());
        if (pid) {
          try {
            process.kill(pid, "SIGKILL");
          } catch {
            // process may have already exited
          }
        }
      }
    }
  } catch {
    // ignore errors from process listing
  }
  await fs.remove(dotnetPrivateInstallPath);
}

export async function withDotnet(
  dotnetChecker: DotnetChecker,
  version: DotnetVersion,
  addToPath: boolean,
  callback: (dotnetExecPath: string) => Promise<void>
): Promise<void> {
  let installDir: string;
  let cleanupCallback: () => void;

  try {
    [installDir, cleanupCallback] = await createTmpDir();
  } catch (error) {
    throw new Error(`Failed to create tmpdir for dotnet, error = '${error}'`);
  }

  const backupPath = process.env.PATH;

  try {
    // use private method as a helper method in test only
    await dotnetChecker["runDotnetInstallScript"](version, installDir);
    const dotnetExecPath = DotnetChecker["getDotnetExecPathFromDotnetInstallationDir"](installDir);

    if (!(await hasDotnetVersion(dotnetExecPath, version))) {
      throw new Error(
        `Failed to install .NET SDK version '${version}' for testing, dotnetExecPath = '${dotnetExecPath}'`
      );
    }

    if (addToPath) {
      process.env.PATH =
        path.resolve(dotnetExecPath, "..") + (isWindows() ? ";" : ":") + process.env.PATH;
    }

    await callback(dotnetExecPath);
  } finally {
    if (addToPath) {
      process.env.PATH = backupPath;
    }
    cleanupCallback();
  }
}

export async function createTmpBackendProjectDir(
  csprojFileName: string
): Promise<[string, () => void]> {
  const [dir, cleanupCallback] = await createTmpDir();

  const csprojPath = path.resolve(
    __dirname,
    "../../../../../../templates/function-base/ts/default/extensions.csproj"
  );
  const targetPath = path.join(dir, csprojFileName);
  await fs.copyFile(csprojPath, targetPath, fs.constants.COPYFILE_EXCL);

  return [dir, cleanupCallback];
}

export async function createMockResourceDir(dirName: string): Promise<[string, () => void]> {
  const [dir, cleanupCallback] = await createTmpDir();

  const resourceDir = path.resolve(__dirname, "../../../../resource/deps-checker");
  const targetDir = path.join(dir, dirName);

  await fs.ensureDir(targetDir);
  await fs.copy(resourceDir, targetDir);

  return [targetDir, cleanupCallback];
}
