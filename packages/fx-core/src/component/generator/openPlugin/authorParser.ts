// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ParsedAuthor } from "./types";

const URL_RE = /^(.+?)\s*\((https?:\/\/[^\s)]+)\)\s*$/;
const EMAIL_RE = /^(.+?)\s*<([^>]+)>\s*$/;

/**
 * Parse the Open Plugin `author` field. The spec allows either an object
 * `{ name, email, url }` or a free-form string like `"Name <email> (url)"`.
 * Returns whatever subset of fields could be extracted.
 */
export function parseAuthor(author: unknown): ParsedAuthor {
  const result: ParsedAuthor = {};
  if (author === null || author === undefined) {
    return result;
  }

  if (typeof author === "string") {
    let trimmed = author.trim();
    const urlMatch = URL_RE.exec(trimmed);
    if (urlMatch) {
      result.url = urlMatch[2];
      trimmed = urlMatch[1].trim();
    }
    const emailMatch = EMAIL_RE.exec(trimmed);
    if (emailMatch) {
      result.email = emailMatch[2].trim();
      trimmed = emailMatch[1].trim();
    }
    if (trimmed) {
      result.name = trimmed;
    }
    return result;
  }

  if (typeof author === "object") {
    const obj = author as { name?: unknown; email?: unknown; url?: unknown };
    if (typeof obj.name === "string" && obj.name) {
      result.name = obj.name;
    }
    if (typeof obj.email === "string" && obj.email) {
      result.email = obj.email;
    }
    if (typeof obj.url === "string" && obj.url) {
      result.url = obj.url;
    }
  }
  return result;
}
