// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import fetch, { Response } from "node-fetch";

export type DownloadOptions = {
  timeout?: number; // 请求超时时间（毫秒）
  maxRedirects?: number; // 最大重定向次数
  progress?: (downloaded: number, total: number) => void; // 进度回调
};

class HttpClient {
  async get(url: string, options: DownloadOptions = {}): Promise<string | Buffer> {
    const { timeout = 30000, progress } = options;
    const res: Response = await fetch(url, {
      redirect: "follow",
      follow: options.maxRedirects ?? 5, // 默认最多跟随 5 次重定向
      timeout,
    });
    if (!res.ok) {
      throw new Error(`Request failed with status ${res.status}`);
    }
    const contentType = res.headers.get("content-type") || "";
    const totalSize = parseInt(res.headers.get("content-length") || "0", 10);
    let downloaded = 0;
    const chunks: Buffer[] = [];
    for await (const chunk of res.body as AsyncIterable<Buffer>) {
      chunks.push(chunk);
      downloaded += chunk.length;
      if (progress) {
        progress(downloaded, totalSize);
      }
    }
    const buffer = Buffer.concat(chunks);
    if (contentType.includes("application/json")) {
      return buffer.toString("utf-8");
    } else if (contentType.includes("text") || contentType.includes("html")) {
      return buffer.toString("utf-8");
    } else {
      return buffer;
    }
  }

  async headTime(url: string, options: DownloadOptions = {}): Promise<number> {
    const { timeout = 30000 } = options;
    const startTime = Date.now();
    const res: Response = await fetch(url, {
      method: "HEAD",
      timeout,
    });
    if (!res.ok) {
      throw new Error(`Request failed with status ${res.status}`);
    }
    return Date.now() - startTime;
  }
}

export const httpClient = new HttpClient();

// async function main() {
//   const url = "https://registry.npmmirror.com/-/binary/node/v22.14.0/node-v22.14.0-win-x64.zip";
//   const result = await httpClient.get(url, {
//     progress: (downloaded, total) => {
//       const progress = ((downloaded / total) * 100).toFixed(2);
//       process.stdout.write(`${progress}%\r`);
//     },
//   });
//   console.log("下载完成", result.length, "字节");
// }
// main();
