// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import fetch, { Response } from "node-fetch";

export type DownloadOptions = {
  timeout?: number;
  maxRedirects?: number;
  progress?: (downloaded: number, total: number) => void;
};

class HttpClient {
  async get(url: string, options: DownloadOptions = {}): Promise<Buffer> {
    const { timeout = 30000, progress } = options;
    const res: Response = await fetch(url, {
      redirect: "follow",
      follow: options.maxRedirects ?? 5,
      timeout,
    });
    if (!res.ok) {
      throw new Error(`Request failed with status ${res.status}`);
    }
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
    return buffer;
  }

  async getText(url: string, options: DownloadOptions = {}): Promise<string> {
    const buffer = await this.get(url, options);
    return buffer.toString("utf-8");
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
