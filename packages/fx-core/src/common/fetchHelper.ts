// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// CJS wrapper around node-fetch v3 (ESM-only).
// Allows sinon to stub fetch in tests, which is not possible when
// directly importing ESM namespace objects (they are frozen/read-only).

import type { RequestInit, Response } from "node-fetch";

export default async function fetch(url: string | URL, init?: RequestInit): Promise<Response> {
  const { default: nodeFetch } = await import("node-fetch");
  return nodeFetch(url, init);
}
