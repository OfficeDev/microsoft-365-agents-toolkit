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

interface NodeDownloadMirror {
  name: string;
  url: string;
  indexJsonUrl: string;
  indexJson?: [
    {
      lts: false | string;
      version: string;
    }
  ];
  html?: string;
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
    const platform = process.platform;
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

    if (targetOS === "darwin" && targetArch === "arm64") {
      return { name: "darwin-arm64", ext: ".tar.xz" };
    }

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

  getLatestLTSVersion(mirror: NodeDownloadMirror): Result<string, InstallNodeJSError> {
    const jsonData = mirror.indexJson!;
    const ltsVersion = jsonData.find(
      (entry: { lts: false | string; version: string }) => entry.lts !== false
    );
    if (!ltsVersion) {
      return err(new InstallNodeJSError("No LTS version found"));
    }
    return ok(ltsVersion.version);
  }

  async fetchJSON(url: string): Promise<Result<any, InstallNodeJSError>> {
    try {
      const res = await httpClient.get(url);
      return ok(JSON.parse(res as string));
    } catch (e: any) {
      return err(
        new InstallNodeJSError(`Failed to fetch JSON from ${url}: ${(e as Error).message}`)
      );
    }
  }

  async fetchString(url: string, timeout?: number): Promise<Result<string, InstallNodeJSError>> {
    try {
      const res = await httpClient.get(url, { timeout: timeout });
      return ok(res as string);
    } catch (e: any) {
      return err(
        new InstallNodeJSError(`Failed to fetch text from ${url}: ${(e as Error).message}`)
      );
    }
  }

  resolveUrl(baseUrl: string, href: string): string | undefined {
    try {
      if (href.startsWith("http")) {
        return href;
      }
      if (href.startsWith("/")) {
        const domain = this.getBaseUrl(baseUrl);
        return new URL(href, domain).href;
      }
      return new URL(href, baseUrl).href;
    } catch (e) {
      return undefined;
    }
  }
  getBaseUrl(url: string): string {
    const parsed = new URL(url);
    return `${parsed.origin}/`; // 取网站根目录
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
      const indexJson = await httpClient.get(mirror.indexJsonUrl, { timeout: timeout });
      const time2 = Date.now();
      mirror.indexJson = JSON.parse(indexJson as string);

      const versionRes = nodejsInstaller.getLatestLTSVersion(mirror);
      if (versionRes.isErr()) {
        return mirror;
      }
      const ltsVersion = versionRes.value;
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
    for (let i = 0; i < 3; ++i) {
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
      return ok(res as Buffer);
    } catch (e: any) {
      return err(
        new InstallNodeJSError(`Failed to fetch binary from ${url}: ${(e as Error).message}`)
      );
    }
  }

  extractZip(buffer: Buffer, targetDir: string): void {
    const zip = new AdmZip(buffer);
    zip.extractAllTo(targetDir, true);
  }

  extractTar(buffer: Buffer, targetDir: string): Result<undefined, InstallNodeJSError> {
    const extname = path.extname(targetDir).toLowerCase();
    if (extname === ".gz" || extname === ".tar.gz") {
      const stream = require("stream");
      const bufferStream = new stream.PassThrough();
      bufferStream.end(buffer);
      bufferStream.pipe(extract({ cwd: targetDir }));
      return ok(undefined);
    } else if (extname === ".xz" || extname === ".tar.xz") {
      const stream = require("stream");
      const bufferStream = new stream.PassThrough();
      bufferStream.end(buffer);
      bufferStream.pipe(extract({ cwd: targetDir }));
      return ok(undefined);
    } else {
      return err(new InstallNodeJSError("Not supported compress file type: " + extname));
    }
  }

  extractPackage(
    buffer: Buffer,
    fileName: string,
    targetDir: string
  ): Result<undefined, InstallNodeJSError> {
    const extname = path.extname(fileName).toLowerCase();
    if (extname === ".zip") {
      this.extractZip(buffer, targetDir);
      return ok(undefined);
    } else if (extname === ".tar.gz" || extname === ".tar.xz") {
      const res = this.extractTar(buffer, targetDir);
      if (res.isErr()) {
        return err(res.error);
      }
      return ok(undefined);
    } else {
      return err(new InstallNodeJSError("Not supported compress file type: " + extname));
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
          new InstallNodeJSError(`Unable to find ${version} in version list page: ${mirror.url}`)
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
        entry.name.includes(osArchName + ext)
      );
      if (!packageUrlEntry) {
        return err(
          new InstallNodeJSError(
            `Unable to find ${osArchName + ext} in package list page: ${versionUrl}`
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
          new InstallNodeJSError(`Unable to find ${version} in version list page: ${mirror.url}`)
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
            `Unable to find ${version}-${osArchName}${ext} in package list page: ${versionUrl}`
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
    const progressBar = context.ui?.createProgressBar("Install NodeJS", 5);
    progressBar?.start();

    // Checking NodeJS in system environment
    context.logProvider?.info("Checking NodeJS in system environment");
    progressBar?.next("Checking NodeJS in system environment");
    if (checkSystemInstalled) {
      const nodeVersion = await NodeChecker.getInstalledNodeVersion();
      if (nodeVersion !== null) {
        context.logProvider?.info(
          `NodeJS is installed in system environment: ${nodeVersion.version}`
        );
        progressBar?.end(true);
        return ok({ status: "ignore" });
      }
    }

    // Checking NodeJS in user folder
    context.logProvider?.info("Checking NodeJS in user folder");
    progressBar?.next("Checking NodeJS in user folder");
    const { name, ext } = this.getNameAndExt();
    const downloadDir = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "nodejs");
    if (checkUserFolderInstalled) {
      const subFolders = await fs.readdir(downloadDir);
      const foundFolder = subFolders.find((subFolder) => subFolder.endsWith(name));
      if (foundFolder) {
        context.logProvider?.info(
          `NodeJS is installed in user folder: ${path.join(downloadDir, foundFolder)}`
        );
        progressBar?.end(true);
        return ok({ status: "ignore", installPath: path.join(downloadDir, foundFolder) });
      } else {
        context.logProvider?.info(`NodeJS not found in user folder: ${downloadDir}`);
      }
    }

    // Testing speed of download mirrors
    context.logProvider?.info("Testing speed of download mirrors");
    progressBar?.next("Testing speed of download mirrors");
    const bestMirror = await nodejsInstaller.getBestMirror(name, ext, context.logProvider);
    if (!bestMirror || !bestMirror.packageUrl) {
      progressBar?.end(true);
      return err(new InstallNodeJSError("All mirrors are not reachable"));
    }
    context.logProvider?.debug(
      `The fastest download mirror is: ${bestMirror.name} - ${bestMirror.url}, latency: ${
        bestMirror.time ?? "unknown"
      } ms`
    );

    // User confirmation for installation
    const confirmRes = await context.ui?.confirm?.({
      name: "confirm",
      title: `Do you want to install NodeJS LTS version (${bestMirror.version!}) in your user folder?`,
    });
    if (confirmRes?.isOk()) {
      if (!confirmRes.value.result) {
        context.logProvider?.warning("User canceled installation");
        progressBar?.end(true);
        return ok({ status: "ignore" });
      }
    } else if (confirmRes?.isErr()) {
      progressBar?.end(true);
      return err(confirmRes.error);
    }

    // Downloading NodeJS package
    context.logProvider?.info(`Downloading binary package: ${bestMirror.packageUrl}`);
    progressBar?.next(`Downloading binary package: ${bestMirror.packageUrl}`);
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
      `Successfully download NodeJS package: ${bestMirror.packageUrl}, size: ${
        binary.length
      }, time: ${t2 - t1} ms`
    );

    // Extracting package
    context.logProvider?.info("Extracting package");
    progressBar?.next("Extracting package");
    await fs.ensureDir(downloadDir);
    const extractRes = nodejsInstaller.extractPackage(binary, bestMirror.packageUrl, downloadDir);
    if (extractRes.isErr()) {
      progressBar?.end(true);
      return err(extractRes.error);
    }
    const targetNodeJSPath = path.join(downloadDir, `node-${bestMirror.version!}-${name}`);
    context.logProvider?.info(
      `Successfully extract NodeJS package in target folder: ${targetNodeJSPath}`
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
