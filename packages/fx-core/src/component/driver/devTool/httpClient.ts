// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import type { Response } from "node-fetch";
import fetch from "../../../common/fetchHelper";

export const httpClientDeps = {
  fetch,
};

export type DownloadOptions = {
  timeout?: number;
  maxRedirects?: number;
  progress?: (downloaded: number, total: number) => void;
};

class HttpClient {
  private createTimeoutSignal(ms: number): { signal: AbortSignal; clear: () => void } {
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), ms);
    return { signal: controller.signal, clear: () => clearTimeout(id) };
  }

  async get(url: string, options: DownloadOptions = {}): Promise<Buffer> {
    const { timeout = 30000, progress } = options;
    const { signal, clear } = this.createTimeoutSignal(timeout);
    try {
      const res: Response = await httpClientDeps.fetch(url, {
        redirect: "follow",
        follow: options.maxRedirects ?? 5,
        signal: signal as any,
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
    } finally {
      clear();
    }
  }

  async getText(url: string, options: DownloadOptions = {}): Promise<string> {
    const buffer = await this.get(url, options);
    return buffer.toString("utf-8");
  }

  async headTime(url: string, options: DownloadOptions = {}): Promise<number> {
    const { timeout = 30000 } = options;
    const { signal, clear } = this.createTimeoutSignal(timeout);
    try {
      const startTime = Date.now();
      const res: Response = await httpClientDeps.fetch(url, {
        method: "HEAD",
        signal: signal as any,
      });
      if (!res.ok) {
        throw new Error(`Request failed with status ${res.status}`);
      }
      return Date.now() - startTime;
    } finally {
      clear();
    }
  }
}

export const httpClient = new HttpClient();
