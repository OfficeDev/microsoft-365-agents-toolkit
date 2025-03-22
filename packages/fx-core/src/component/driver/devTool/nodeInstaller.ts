// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ConfigFolderName, err, LogProvider, ok, Result } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import { parseHTML } from "linkedom";
import os from "os";
import * as path from "path";
import { extract } from "tar";
import { InstallNodeJSError } from "../../../error";
import { NodeChecker } from "../../deps-checker/internal/nodeChecker";
import { WrapDriverContext } from "../util/wrapUtil";
import { httpClient } from "./httpClient";
import { getLocalizedString } from "../../../common/localizeUtils";
import stream from "stream";

export const ComponentName = "NodeInstaller";

export interface NodeDownloadMirror {
  name: string;
  url: string;
  indexJsonUrl: string;
  indexJson?: {
    lts: false | string;
    version: string;
  }[];
  version?: string;
  versionUrl?: string;
  packageUrl?: string;
  time?: number;
}
const FirstPriorityMirror: NodeDownloadMirror = {
  name: "NPM",
  url: "https://registry.npmmirror.com/-/binary/node/",
  indexJsonUrl: "https://cdn.npmmirror.com/binaries/node/index.json",
};
// const FirstPriorityMirror: NodeDownloadMirror = {
//   name: "Official",
//   url: "https://nodejs.org/dist/",
//   indexJsonUrl: "https://nodejs.org/dist/index.json",
// };
const BackupMirrors: NodeDownloadMirror[] = [
  {
    name: "Official",
    url: "https://nodejs.org/dist/",
    indexJsonUrl: "https://nodejs.org/dist/index.json",
  },
  {
    name: "Tencent",
    url: "https://mirrors.cloud.tencent.com/nodejs-release/",
    indexJsonUrl: "https://mirrors.cloud.tencent.com/nodejs-release/index.json",
  },
  {
    name: "Aliyun",
    url: "https://mirrors.aliyun.com/nodejs-release/",
    indexJsonUrl: "https://mirrors.aliyun.com/nodejs-release/index.json",
  },
];

export interface EnsureNodeJSResult {
  status: "ignore" | "installed";
  installPath?: string;
  totalTime?: number;
}

export class NodejsInstaller {
  getNameAndExt(): { name: string; ext: string } {
    const platform = os.platform();
    const arch = os.arch();

    const osMap: { [key: string]: string } = {
      win32: "win",
      darwin: "darwin",
      linux: "linux",
      aix: "aix",
    };

    const archMap: { [key: string]: string } = {
      x64: "x64",
      arm64: "arm64",
      arm: "armv7l",
      ppc64: "ppc64le",
      s390x: "s390x",
    };

    const targetOS = osMap[platform] || platform;
    const targetArch = archMap[arch] || arch;

    let extPattern;
    switch (targetOS) {
      case "win":
        extPattern = ".zip"; // Windows zip is preferred
        break;
      case "darwin":
      case "linux":
        extPattern = ".tar.xz"; // macOS/Linux .tar.xz is preferred
        break;
      default:
        extPattern = ".tar.gz";
    }
    return { name: `${targetOS}-${targetArch}`, ext: extPattern };
  }

  getLatestLTSVersion(mirror: NodeDownloadMirror): string | undefined {
    const jsonData = mirror.indexJson!;
    const ltsVersion = jsonData.find(
      (entry: { lts: false | string; version: string }) => entry.lts !== false
    );
    return ltsVersion?.version;
  }

  async fetchJSON(url: string): Promise<Result<any, InstallNodeJSError>> {
    try {
      const res = await httpClient.getText(url);
      return ok(JSON.parse(res));
    } catch (e: any) {
      return err(new InstallNodeJSError((e as Error).message));
    }
  }

  async fetchString(url: string, timeout?: number): Promise<Result<string, InstallNodeJSError>> {
    try {
      const res = await httpClient.getText(url, { timeout: timeout });
      return ok(res);
    } catch (e: any) {
      return err(new InstallNodeJSError((e as Error).message));
    }
  }

  resolveUrl(baseUrl: string, href: string): string {
    return new URL(href, baseUrl).toString();
  }

  async testMirrorSpeed(
    mirror: NodeDownloadMirror,
    osArchName: string,
    ext: string,
    timeout: number,
    logger?: LogProvider
  ): Promise<NodeDownloadMirror> {
    try {
      const headTime = await httpClient.headTime(mirror.url, { timeout: timeout });
      const time1 = Date.now();
      const indexJson = await httpClient.getText(mirror.indexJsonUrl, { timeout: timeout });
      const time2 = Date.now();
      mirror.indexJson = JSON.parse(indexJson);

      const ltsVersion = nodejsInstaller.getLatestLTSVersion(mirror);
      if (!ltsVersion) {
        return mirror;
      }
      mirror.version = ltsVersion;
      const packageUrlRes = await this.getDownloadUrl(mirror, ltsVersion, osArchName, ext);
      if (packageUrlRes.isErr()) {
        return mirror;
      }
      const packageUrl = packageUrlRes.value;
      const time3 = Date.now();
      mirror.packageUrl = packageUrl;
      mirror.time = headTime + time3 - time1;
      logger?.debug(
        `Mirror: ${mirror.name}, URL: ${mirror.url}, Head: ${headTime} ms, Index JSON: ${
          time2 - time1
        } ms, Get URL: ${time3 - time2} ms, Total: ${mirror.time} ms`
      );
    } catch (e: any) {
      logger?.error(`Mirror: ${mirror.name}, URL: ${mirror.url}, Error: ${(e as Error).message}`);
    }
    return mirror;
  }

  async getBestMirror(
    osArchName: string,
    ext: string,
    logger?: LogProvider
  ): Promise<NodeDownloadMirror | undefined> {
    const mirror = await this.testMirrorSpeed(FirstPriorityMirror, osArchName, ext, 1000, logger);
    if (mirror.packageUrl) {
      return mirror;
    }
    for (let i = 0; i < 5; ++i) {
      const mirror = await Promise.race(
        BackupMirrors.map((mirror) => this.testMirrorSpeed(mirror, osArchName, ext, 1000, logger))
      );
      if (mirror.packageUrl) {
        return mirror;
      }
    }
    return undefined;
  }

  parseHtmlToGetUrl(url: string, html: string, pattern: string): string | undefined {
    const { document } = parseHTML(html);
    const links = [...document.querySelectorAll("a")]
      .map((a) => a.getAttribute("href"))
      .filter((href) => href && href.includes(pattern))
      .map((href) => this.resolveUrl(url, href!));
    if (links.length === 0) {
      return undefined;
    }
    return links[0];
  }

  async fetchBinary(
    url: string,
    timeout?: number,
    onProgress?: (progress: string) => void
  ): Promise<Result<Buffer, InstallNodeJSError>> {
    try {
      const res = await httpClient.get(url, {
        timeout: timeout,
        progress: (downloaded, total) => {
          if (onProgress) {
            const progress = ((downloaded / total) * 100).toFixed(2);
            onProgress(`download progress: ${progress}%`);
          }
        },
      });
      return ok(res);
    } catch (e: any) {
      return err(new InstallNodeJSError((e as Error).message));
    }
  }

  getAdmZip(buffer: Buffer): AdmZip {
    return new AdmZip(buffer);
  }

  extractZip(buffer: Buffer, targetDir: string): void {
    const zip = this.getAdmZip(buffer);
    zip.extractAllTo(targetDir, true);
  }

  extractTar(buffer: Buffer, fileName: string, targetDir: string): void {
    const extname = path.extname(fileName).toLowerCase();
    if (extname === ".gz" || extname === ".tar.gz") {
      const bufferStream = new stream.PassThrough();
      bufferStream.end(buffer);
      bufferStream.pipe(extract({ cwd: targetDir }));
    } else if (extname === ".xz" || extname === ".tar.xz") {
      const bufferStream = new stream.PassThrough();
      bufferStream.end(buffer);
      bufferStream.pipe(extract({ cwd: targetDir }));
    }
  }

  extractPackage(buffer: Buffer, fileName: string, targetDir: string): void {
    if (fileName.endsWith(".zip")) {
      this.extractZip(buffer, targetDir);
    } else if (fileName.endsWith(".tar.gz") || fileName.endsWith(".tar.xz")) {
      this.extractTar(buffer, fileName, targetDir);
    }
  }

  async getDownloadUrl(
    mirror: NodeDownloadMirror,
    version: string,
    osArchName: string,
    ext: string
  ): Promise<Result<string, InstallNodeJSError>> {
    if (mirror.name === "NPM") {
      const jsonRes = await this.fetchJSON(mirror.url);
      if (jsonRes.isErr()) {
        return err(jsonRes.error);
      }
      const json = jsonRes.value as { url: string; name: string }[];
      const versionInfo = json.find((entry: { url: string; name: string }) =>
        entry.name.includes(version)
      );
      if (!versionInfo) {
        return err(
          new InstallNodeJSError(
            getLocalizedString("action.devTool.nodeInstaller.UnableToFind", version, mirror.url)
          )
        );
      }
      const versionUrl = versionInfo.url;
      mirror.versionUrl = versionUrl;
      const versionJsonRes = await this.fetchJSON(versionUrl);
      if (versionJsonRes.isErr()) {
        return err(versionJsonRes.error);
      }
      const versionJson = versionJsonRes.value;
      const packageUrlEntry = versionJson.find((entry: any) =>
        entry.name.includes(`${version}-${osArchName}${ext}`)
      );
      if (!packageUrlEntry) {
        return err(
          new InstallNodeJSError(
            getLocalizedString(
              "action.devTool.nodeInstaller.UnableToFind",
              `${version}-${osArchName}${ext}`,
              versionUrl
            )
          )
        );
      }
      return ok(packageUrlEntry.url);
    } else {
      const versionListHtmlRes = await this.fetchString(mirror.url);
      if (versionListHtmlRes.isErr()) {
        return err(versionListHtmlRes.error);
      }
      const versionListHtml = versionListHtmlRes.value;
      const versionUrl = nodejsInstaller.parseHtmlToGetUrl(mirror.url, versionListHtml, version);
      if (!versionUrl) {
        return err(
          new InstallNodeJSError(
            getLocalizedString("action.devTool.nodeInstaller.UnableToFind", version, mirror.url)
          )
        );
      }
      mirror.versionUrl = versionUrl;
      const packageListHtmlRes = await this.fetchString(versionUrl);
      if (packageListHtmlRes.isErr()) {
        return err(packageListHtmlRes.error);
      }
      const packageListHtml = packageListHtmlRes.value;
      const packageUrl = nodejsInstaller.parseHtmlToGetUrl(
        versionUrl,
        packageListHtml,
        `${version}-${osArchName}${ext}`
      );
      if (!packageUrl) {
        return err(
          new InstallNodeJSError(
            getLocalizedString(
              "action.devTool.nodeInstaller.UnableToFind",
              `${version}-${osArchName}${ext}`,
              versionUrl
            )
          )
        );
      }
      return ok(packageUrl);
    }
  }

  async ensureNodeJS(
    context: WrapDriverContext,
    checkSystemInstalled = true,
    checkUserFolderInstalled = false
  ): Promise<Result<EnsureNodeJSResult, InstallNodeJSError>> {
    const startTime = Date.now();
    const progressBar = context.ui?.createProgressBar(
      getLocalizedString("action.devTool.nodeInstaller.Progress.title"),
      5
    );
    progressBar?.start();

    // Checking NodeJS in system environment
    let progressText = getLocalizedString("action.devTool.nodeInstaller.Progress1");
    context.logProvider?.info(progressText);
    progressBar?.next(progressText);
    if (checkSystemInstalled) {
      const nodeVersion = await NodeChecker.getInstalledNodeVersion();
      if (nodeVersion !== null) {
        context.logProvider?.info(
          getLocalizedString("action.devTool.nodeInstaller.InstalledSystem", nodeVersion.version)
        );
        progressBar?.end(true);
        return ok({ status: "ignore" });
      } else {
        context.logProvider?.info(
          getLocalizedString("action.devTool.nodeInstaller.NotInstalledSystem")
        );
      }
    }

    // Checking NodeJS in user folder
    progressText = getLocalizedString("action.devTool.nodeInstaller.Progress2");
    context.logProvider?.info(progressText);
    progressBar?.next(progressText);
    const { name, ext } = this.getNameAndExt();
    const downloadDir = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "nodejs");
    if (checkUserFolderInstalled) {
      const subFolders = await fs.readdir(downloadDir);
      const foundFolder = subFolders.find((subFolder) => subFolder.endsWith(name));
      if (foundFolder) {
        context.logProvider?.info(
          getLocalizedString(
            "action.devTool.nodeInstaller.InstalledUser",
            path.join(downloadDir, foundFolder)
          )
        );
        progressBar?.end(true);
        return ok({ status: "ignore", installPath: path.join(downloadDir, foundFolder) });
      } else {
        context.logProvider?.info(
          getLocalizedString("action.devTool.nodeInstaller.NotInstalledUser", downloadDir)
        );
      }
    }

    // Testing speed of download mirrors
    progressText = getLocalizedString("action.devTool.nodeInstaller.Progress3");
    context.logProvider?.info(progressText);
    progressBar?.next(progressText);
    const bestMirror = await nodejsInstaller.getBestMirror(name, ext, context.logProvider);
    if (!bestMirror?.packageUrl) {
      progressBar?.end(true);
      return err(
        new InstallNodeJSError(getLocalizedString("action.devTool.nodeInstaller.NoMirrorUsable"))
      );
    }
    context.logProvider?.info(
      getLocalizedString(
        "action.devTool.nodeInstaller.BestMirror",
        bestMirror.name,
        bestMirror.url,
        bestMirror.time
      )
    );

    // User confirmation for installation
    const confirmRes = await context.ui?.confirm?.({
      name: "confirm",
      title: getLocalizedString("action.devTool.nodeInstaller.Confirm", bestMirror.version!),
    });

    if (confirmRes?.isErr()) {
      progressBar?.end(true);
      return err(confirmRes.error);
    }

    // Downloading NodeJS package
    progressText = getLocalizedString(
      "action.devTool.nodeInstaller.Progress4",
      bestMirror.packageUrl
    );
    context.logProvider?.info(progressText);
    progressBar?.next(progressText);
    const t1 = Date.now();
    const packageRes = await nodejsInstaller.fetchBinary(
      bestMirror.packageUrl,
      undefined,
      (progress) => void progressBar?.text?.(progress)
    );
    const t2 = Date.now();
    if (packageRes.isErr()) {
      progressBar?.end(true);
      return err(packageRes.error);
    }
    const binary = packageRes.value;
    context.logProvider?.info(
      getLocalizedString(
        "action.devTool.nodeInstaller.SuccessDownload",
        bestMirror.packageUrl,
        binary.length,
        t2 - t1
      )
    );

    // Extracting package
    progressText = getLocalizedString("action.devTool.nodeInstaller.Progress5");
    context.logProvider?.info(progressText);
    progressBar?.next(progressText);
    await fs.ensureDir(downloadDir);
    nodejsInstaller.extractPackage(binary, bestMirror.packageUrl, downloadDir);
    const targetNodeJSPath = path.join(downloadDir, `node-${bestMirror.version!}-${name}`);
    context.logProvider?.info(
      getLocalizedString("action.devTool.nodeInstaller.SuccessExtract", targetNodeJSPath)
    );
    progressBar?.end(true);
    const totalTime = Date.now() - startTime;
    return ok({ status: "installed", installPath: targetNodeJSPath, totalTime: totalTime });
  }
}

export const nodejsInstaller = new NodejsInstaller();

// const logProvider = {
//   info: (message: string) => {
//     console.log(message);
//   },
//   debug: (message: string) => {
//     console.debug(message);
//   },
//   error: (message: string) => {
//     console.error(message);
//   },
// };
// async function main() {
//   const result = await nodejsInstaller.ensureNodeJS(
//     {
//       logProvider: logProvider,
//       ui: {
//         createProgressBar: (title: string, totalSteps: number) => {
//           return {
//             start: (message?: string) => {
//               console.log(`${title}: ${message || ""}`);
//             },
//             next: (message?: string) => {
//               console.log(`Next step: ${message || ""}`);
//             },
//             end: (success: boolean) => {
//               console.log(`Progress ended. Success: ${success.toString()}`);
//             },
//             text: (message: string) => {
//               process.stdout.write(`Progress: ${message}\r`);
//             },
//           };
//         },
//       },
//     } as any,
//     false,
//     false
//   );
//   console.log(result);
// }
// main();
