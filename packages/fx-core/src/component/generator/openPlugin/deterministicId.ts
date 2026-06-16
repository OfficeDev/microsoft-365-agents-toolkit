// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createHash } from "crypto";

// RFC 4122 URL namespace UUID.
const URL_NAMESPACE_BYTES = Buffer.from([
  0x6b, 0xa7, 0xb8, 0x11, 0x9d, 0xad, 0x11, 0xd1, 0x80, 0xb4, 0x00, 0xc0, 0x4f, 0xd4, 0x30, 0xc8,
]);

/**
 * Compute a deterministic UUIDv5 (SHA-1 based) of `seed` under the URL
 * namespace. The same `seed` always produces the same UUID.
 */
export function deterministicAppId(seed: string): string {
  const nameBytes = Buffer.from(seed, "utf8");
  const hash = createHash("sha1").update(URL_NAMESPACE_BYTES).update(nameBytes).digest();
  const bytes = Buffer.from(hash.subarray(0, 16));
  // version (5) and variant bits per RFC 4122
  bytes[6] = (bytes[6] & 0x0f) | 0x50;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = bytes.toString("hex");
  return [
    hex.substring(0, 8),
    hex.substring(8, 12),
    hex.substring(12, 16),
    hex.substring(16, 20),
    hex.substring(20),
  ].join("-");
}
