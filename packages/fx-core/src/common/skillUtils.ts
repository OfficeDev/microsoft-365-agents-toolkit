// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export function isValidSkillName(name: unknown): name is string {
  return typeof name === "string" && name.length <= 64 && /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name);
}
