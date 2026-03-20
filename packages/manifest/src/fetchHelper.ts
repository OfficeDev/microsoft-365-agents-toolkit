// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// CJS wrapper around node-fetch v3 (ESM-only).
// Allows sinon to stub fetch in tests, which is not possible when
// directly importing ESM namespace objects (they are frozen/read-only).
//
// Uses new Function() to issue a real ESM dynamic import() that TypeScript
// will not transform into require() under "module": "commonjs".

// eslint-disable-next-line @typescript-eslint/no-implied-eval,no-new-func
const _importDynamic = new Function("modulePath", "return import(modulePath)");

export default async function fetch(
  url: string | URL,
  init?: import("node-fetch").RequestInit
): Promise<import("node-fetch").Response> {
  const { default: nodeFetch } = await _importDynamic("node-fetch");
  return nodeFetch(url, init);
}
