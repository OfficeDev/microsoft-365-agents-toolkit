// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";

/**
 * Resolve the root templates folder shipped alongside core-next.
 * At runtime `__dirname` is `build/` so `../templates` resolves to
 * `packages/core-next/templates/`.
 */
export function getTemplatesFolder(): string {
  return path.resolve(__dirname, "../templates");
}
