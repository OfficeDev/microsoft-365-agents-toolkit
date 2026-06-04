// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Fetch wrapper that uses the built-in global fetch (Node 18+) when available,
// falling back to node-fetch v3 (ESM-only) for older runtimes.
// The CJS wrapper with new Function() avoids TypeScript transforming
// import() into require() under "module": "commonjs".
// This module is also stubbable via sinon, unlike frozen ESM namespace objects.

import type { RequestInit, Response } from "node-fetch";

let nodeFetchPromise:
  | Promise<{ default: (url: string, init?: RequestInit) => Promise<Response> }>
  | undefined;

async function loadNodeFetch(): Promise<{
  default: (url: string, init?: RequestInit) => Promise<Response>;
}> {
  if (!nodeFetchPromise) {
    nodeFetchPromise = import("node-fetch") as any;
  }
  return nodeFetchPromise!;
}

export default async function fetch(url: string | URL, init?: RequestInit): Promise<Response> {
  const urlString = url instanceof URL ? url.href : url;
  if (!/^https?:\/\//i.test(urlString)) {
    throw new Error(`fetch only supports HTTP/HTTPS URLs, got: ${urlString}`);
  }
  if (typeof globalThis.fetch === "function") {
    return globalThis.fetch(urlString, init as any) as any;
  }
  const { default: nodeFetch } = await loadNodeFetch();
  return nodeFetch(urlString, init);
}
