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

import { httpClient } from "./httpClient";
import { WrapDriverContext } from "../util/wrapUtil";
import { createSymlink } from "../../deps-checker/util/fileHelper";

interface NodeDownloadMirror {
  name: string;
  url: string;
  html?: string;
  time?: number;
}

export interface EnsureNodeJSResult {
  status: "ignore" | "installed";
  installPath?: string;
}

export class NodejsInstaller {
  getCompatibleNodeJSVersionSuffix(): { name: string; ext: string } {
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

  async getLatestLTSVersion(): Promise<Result<string, InstallNodeJSError>> {
    const fetchRes = await this.fetchJSON("https://nodejs.org/dist/index.json");
    if (fetchRes.isErr()) {
      return err(fetchRes.error);
    }
    const jsonData = fetchRes.value;
    const ltsVersion = jsonData.find((entry: any) => entry.lts !== false);
    if (!ltsVersion) {
      return err(new InstallNodeJSError("No LTS version found"));
    }
    return ok(ltsVersion.version);
  }

  async getTargetVersionPackageUrl(
    baseUrl: string,
    versionPattern: string
  ): Promise<Result<string, InstallNodeJSError>> {
    const htmlRes = await this.fetchString(baseUrl);
    if (htmlRes.isErr()) {
      return err(htmlRes.error);
    }
    const html = htmlRes.value;
    const { document } = parseHTML(html);
    const compatibleLinks = [...document.querySelectorAll("a")]
      .map((a) => a.getAttribute("href"))
      .filter((href) => href && href.includes("node-") && href.endsWith(versionPattern))
      .map((href) => this.resolveUrl(baseUrl, href!)) as string[];
    if (compatibleLinks.length === 0) {
      err(
        new InstallNodeJSError(
          `No compatible links found for pattern: ${versionPattern} in NodeJS download page: ${baseUrl}`
        )
      );
    }
    const downloadUrl = compatibleLinks[0];
    if (!downloadUrl) {
      return err(
        new InstallNodeJSError(
          `Failed to get download URL for pattern: ${versionPattern} in NodeJS download page: ${baseUrl}`
        )
      );
    }
    return ok(downloadUrl);
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

  async getFastestNodeDownloadMirror(): Promise<NodeDownloadMirror> {
    const mirrors = [
      { name: "Official", url: "https://nodejs.org/dist/" },
      { name: "NPM Mirror", url: "https://registry.npmmirror.com/-/binary/node/" },
      { name: "Tencent", url: "https://mirrors.cloud.tencent.com/nodejs-release/" },
      { name: "Aliyun", url: "https://mirrors.aliyun.com/nodejs-release/" },
    ];
    const results = await Promise.all(
      mirrors.map(async (mirror) => {
        const start = Date.now();
        const fetchRes = await this.fetchString(mirror.url, 1000);
        if (fetchRes.isErr()) {
          return { name: mirror.name, url: mirror.url, time: Infinity, html: undefined };
        }
        const html = fetchRes.value;
        const time = Date.now() - start;
        return { name: mirror.name, url: mirror.url, html: html, time: time };
      })
    );
    const fastest = results.reduce((a, b) => (a.time < b.time ? a : b));
    return fastest;
  }

  parseTargetVersionUrl(url: string, html: string, version: string): string | undefined {
    const { document } = parseHTML(html);
    const links = [...document.querySelectorAll("a")]
      .map((a) => a.getAttribute("href"))
      .filter((href) => href && href.includes(version))
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

  extractZipFromBuffer(buffer: Buffer, targetDir: string): void {
    const zip = new AdmZip(buffer);
    zip.extractAllTo(targetDir, true);
  }

  extractTarFromBuffer(buffer: Buffer, targetDir: string): Result<undefined, InstallNodeJSError> {
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

  extractNodeJsPackageFromBuffer(
    buffer: Buffer,
    fileName: string,
    targetDir: string
  ): Result<undefined, InstallNodeJSError> {
    const extname = path.extname(fileName).toLowerCase();
    if (extname === ".zip") {
      this.extractZipFromBuffer(buffer, targetDir);
      return ok(undefined);
    } else if (extname === ".tar.gz" || extname === ".tar.xz") {
      const res = this.extractTarFromBuffer(buffer, targetDir);
      if (res.isErr()) {
        return err(res.error);
      }
      return ok(undefined);
    } else {
      return err(new InstallNodeJSError("Not supported compress file type: " + extname));
    }
  }

  async ensureNodeJS(
    context: WrapDriverContext,
    checkSystemInstalled = true,
    checkUserFolderInstalled = false
  ): Promise<Result<EnsureNodeJSResult, InstallNodeJSError>> {
    const progressBar = context.ui?.createProgressBar("Install NodeJS", 9);
    progressBar?.start();
    progressBar?.next("Checking NodeJS in system environment.");
    if (checkSystemInstalled) {
      const nodeVersion = await NodeChecker.getInstalledNodeVersion();
      if (nodeVersion !== null) {
        context.logProvider?.debug(`Found NodeJS version: ${nodeVersion.version}`);
        progressBar?.end(true);
        return ok({ status: "ignore" });
      }
    }
    // get the latest LTS version
    progressBar?.next("Looking up latest LTS version of NodeJS");
    const latestVersionRes = await nodejsInstaller.getLatestLTSVersion();
    if (latestVersionRes.isErr()) {
      progressBar?.end(true);
      return err(latestVersionRes.error);
    }
    const latestVersion = latestVersionRes.value;

    progressBar?.next("Checking OS type and architecture");
    const { name, ext } = this.getCompatibleNodeJSVersionSuffix();

    const latestVersionFullname = `node-${latestVersion}-${name}`;

    context.logProvider?.debug(`Latest NodeJS LTS version full name is: ${latestVersionFullname}`);

    const downloadDir = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "nodejs");

    progressBar?.next("Checking compatible NodeJS version in user folder");

    if (checkUserFolderInstalled) {
      const subFolders = await fs.readdir(downloadDir);

      const foundFolder = subFolders.find((subFolder) => subFolder.endsWith(name));

      if (foundFolder) {
        context.logProvider?.debug(
          `Found NodeJS version in user folder: ${path.join(downloadDir, foundFolder)}`
        );
        progressBar?.end(true);
        return ok({ status: "ignore", installPath: path.join(downloadDir, foundFolder) });
      } else {
        context.logProvider?.debug(
          `NodeJS not found in user folder: ${downloadDir}, continue to install`
        );
      }
    }

    const confirmRes = await context.ui?.confirm?.({
      name: "confirm",
      title: "NodeJS is not installed, do you want to install latest LTS version?",
    });

    if (confirmRes?.isOk()) {
      if (!confirmRes.value.result) {
        context.logProvider?.warning("User canceled NodeJS installation");
        progressBar?.end(true);
        return ok({ status: "ignore" });
      }
    } else if (confirmRes?.isErr()) {
      progressBar?.end(true);
      return err(confirmRes.error);
    }

    // get the fastest download mirror
    progressBar?.next("Searching best NodeJS download mirrors");
    const fast = await nodejsInstaller.getFastestNodeDownloadMirror();
    if (!fast.html) {
      progressBar?.end(true);
      return err(new InstallNodeJSError("Failed to get fastest download mirror"));
    }
    context.logProvider?.debug(
      `The fastest NodeJS download mirror site is: ${fast.url}, latency: ${
        fast.time ?? "unknown"
      } ms`
    );

    let versionPackageUrl = "";
    progressBar?.next("Fetching NodeJS download url for version: " + latestVersion);
    if (fast.name === "NPM Mirror") {
      const jsonRes = await this.fetchJSON(fast.url);
      if (jsonRes.isErr()) {
        progressBar?.end(true);
        return err(jsonRes.error);
      }
      const json = jsonRes.value;
      const versionInfo = json.find((entry: any) => entry.name.includes(latestVersion));
      if (!versionInfo) {
        progressBar?.end(true);
        return err(
          new InstallNodeJSError(
            `No compatible links found for pattern: ${latestVersion} in NodeJS download page: ${fast.url}`
          )
        );
      }

      const versionUrl = versionInfo.url;
      const versionJsonRes = await this.fetchJSON(versionUrl);
      if (versionJsonRes.isErr()) {
        progressBar?.end(true);
        return err(versionJsonRes.error);
      }
      const versionJson = versionJsonRes.value;
      const compatibleVersionObj = versionJson.find((entry: any) =>
        entry.name.includes(name + ext)
      );
      if (!compatibleVersionObj) {
        progressBar?.end(true);
        return err(
          new InstallNodeJSError(
            `No compatible links found for pattern: ${name + ext} in NodeJS download page: ${
              fast.url
            }`
          )
        );
      }
      versionPackageUrl = compatibleVersionObj.url;
    } else {
      const versionFolderUrl = nodejsInstaller.parseTargetVersionUrl(
        fast.url,
        fast.html,
        latestVersion
      );
      if (!versionFolderUrl) {
        progressBar?.end(true);
        return err(
          new InstallNodeJSError(
            `Failed to get folder URL for target version: ${latestVersion} in page: ${fast.url}`
          )
        );
      }

      context.logProvider?.debug(`NodeJS download folder page url: ${versionFolderUrl}`);
      const versionPackageUrlRes = await nodejsInstaller.getTargetVersionPackageUrl(
        versionFolderUrl,
        name + ext
      );
      if (versionPackageUrlRes.isErr()) {
        progressBar?.end(true);
        return err(versionPackageUrlRes.error);
      }
      versionPackageUrl = versionPackageUrlRes.value;
    }

    context.logProvider?.debug(`Start to download NodeJS package: ${versionPackageUrl}`);

    progressBar?.next("Fetching NodeJS binary package: " + versionPackageUrl);

    const t1 = Date.now();
    const packageRes = await nodejsInstaller.fetchBinary(versionPackageUrl, undefined, (progress) =>
      process.stdout.write(progress + "\r")
    );
    const t2 = Date.now();

    if (packageRes.isErr()) {
      progressBar?.end(true);
      return err(packageRes.error);
    }

    const binary = packageRes.value;

    context.logProvider?.debug(
      `Successfully download NodeJS package: ${versionPackageUrl}, size: ${binary.length}, time: ${
        t2 - t1
      } ms`
    );

    progressBar?.next("Extracting package");

    await fs.ensureDir(downloadDir);

    const extractRes = nodejsInstaller.extractNodeJsPackageFromBuffer(
      binary,
      versionPackageUrl,
      downloadDir
    );
    if (extractRes.isErr()) {
      progressBar?.end(true);
      return err(extractRes.error);
    }
    context.logProvider?.debug(
      `Successfully extract NodeJS package in target folder: ${downloadDir}`
    );

    const targetNodeJSPath = path.join(downloadDir, latestVersionFullname);

    progressBar?.next("NodeJS installation completed");

    progressBar?.end(true);

    return ok({ status: "installed", installPath: targetNodeJSPath });
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
//     false,
//     false,
//     logProvider as any,
//     (progress) => {
//       console.log("--------------------------" + progress);
//     }
//   );
//   console.log(result);
// }

// main();
