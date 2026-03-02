// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export function parseChallenges(header: string) {
  const schemeSeparator = header.indexOf(" ");
  const challenges = header.substring(schemeSeparator + 1).split(",");
  const challengeMap: { [key: string]: string } = {};

  challenges.forEach((challenge) => {
    const [key, value] = challenge.split("=");
    challengeMap[key.trim()] = decodeURI(value.replace(/['"]+/g, ""));
  });
  return challengeMap;
}

export function decodeClaimsChallenge(encodedClaims: string): string | undefined {
  try {
    return Buffer.from(encodedClaims, "base64").toString("utf8");
  } catch (e) {}
  return undefined;
}
