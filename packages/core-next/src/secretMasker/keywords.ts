// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Credential-related keyword suffixes.
 * A key is considered sensitive if it ends with (case-insensitive) any of these.
 * Derived from fx-core's combinations1/2 expansions — using underscore, hyphen,
 * and camelCase variants of common credential patterns.
 */
const BASE_SUFFIXES = [
  "password",
  "pwd",
  "secret",
  "secretkey",
  "secretvalue",
  "accesskey",
  "apikey",
  "connectionstring",
  "token",
  "accesstoken",
  "refreshtoken",
  "bearertoken",
  "credential",
  "credentials",
  "privatekey",
  "certificate",
  "certificatekey",
  "masterkey",
  "primarykey",
  "secondarykey",
  "sharedaccesskey",
  "storageaccountkey",
  "encryptionkey",
  "signingkey",
  "authorizationheader",
  "authkey",
  "authtoken",
  "clientsecret",
  "clientid",
  "accountkey",
  "instrumentationkey",
  "sastoken",
  "saskey",
  "subscriptionkey",
  "managementkey",
  "appkey",
  "appsecret",
  "webhooksecret",
  "passphrase",
  "keyfile",
  "pemfile",
  "pfxfile",
  "sshkey",
  "gpgkey",
  "symmetrickey",
  "decryptionkey",
  "validationkey",
  "oauthtoken",
  "personalaccesstoken",
  "pat",
];

/**
 * Expand base suffixes into underscore_case and hyphen-case variants.
 */
function expandSuffix(s: string): string[] {
  const variants = [s];
  // Insert underscores before each uppercase letter, e.g. "secretKey" → "secret_key"
  const underscored = s.replace(/([a-z])([A-Z])/g, "$1_$2").toLowerCase();
  if (underscored !== s.toLowerCase()) {
    variants.push(underscored);
  }
  // Hyphen variant
  const hyphenated = underscored.replace(/_/g, "-");
  if (hyphenated !== underscored) {
    variants.push(hyphenated);
  }
  return variants;
}

/**
 * All credential keyword suffixes (case-insensitive matching).
 */
export const CREDENTIAL_KEYWORDS: readonly string[] = Object.freeze(
  BASE_SUFFIXES.flatMap(expandSuffix)
);

/**
 * Check whether a given key name matches a credential keyword (case-insensitive).
 * Returns true if the key ends with any credential keyword suffix.
 */
export function matchesCredentialKeyword(key: string): boolean {
  const lower = key.toLowerCase();
  return CREDENTIAL_KEYWORDS.some((kw) => lower.endsWith(kw));
}
