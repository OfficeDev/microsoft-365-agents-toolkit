// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Mustache from "mustache";

export function renderFragment(lines: readonly string[], values: object): string {
  return Mustache.render(lines.join("\n"), values, undefined, { escape: (value) => value });
}
