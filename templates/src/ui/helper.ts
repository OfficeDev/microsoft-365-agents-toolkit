// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as stringMap from "./resource/package.nls.json";

export function getString(key: string): string {
  return key in stringMap ? (stringMap as Record<string, string>)[key] : key;
}
