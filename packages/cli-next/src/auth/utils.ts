// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AccountInfo } from "@azure/msal-node";

/**
 * Find an account by its homeAccountId in a list of accounts.
 */
export function getAccountByHomeId(
  homeAccountId: string,
  allAccounts: AccountInfo[]
): AccountInfo | null {
  if (homeAccountId && allAccounts && allAccounts.length) {
    return (
      allAccounts.filter((accountObj) => accountObj.homeAccountId === homeAccountId)[0] || null
    );
  }
  return null;
}

/**
 * Decode a JWT token payload section.
 */
export function convertTokenToJson(token: string): Record<string, unknown> {
  const parts = token.split(".");
  // JWE tokens have 5 parts — return empty object
  if (parts.length === 5) return {};
  const buff = Buffer.from(parts[1], "base64");
  return JSON.parse(buff.toString("utf8"));
}

/**
 * Parse WWW-Authenticate header challenges into a map.
 */
export function parseChallenges(header: string): Record<string, string> {
  const schemeSeparator = header.indexOf(" ");
  const challenges = header.substring(schemeSeparator + 1).split(",");
  const challengeMap: Record<string, string> = {};
  for (const challenge of challenges) {
    const [key, value] = challenge.split("=");
    challengeMap[key.trim()] = decodeURI(value.replace(/['"]+/g, ""));
  }
  return challengeMap;
}

/**
 * Decode a base64-encoded claims challenge.
 */
export function decodeClaimsChallenge(encodedClaims: string): string | undefined {
  try {
    return Buffer.from(encodedClaims, "base64").toString("utf8");
  } catch {
    return undefined;
  }
}

/**
 * Check if the machine can reach login.microsoftonline.com.
 */
export function checkIsOnline(): Promise<boolean> {
  return new Promise((resolve) => {
    // Use dynamic import to avoid top-level require
    import("http").then(({ default: http }) => {
      const req = http.request(
        { hostname: "login.microsoftonline.com", path: "/", method: "HEAD" },
        (res) => {
          res.on("data", () => {});
          res.on("end", () => resolve(true));
        }
      );
      req.on("error", () => resolve(false));
      req.end();
    });
  });
}
