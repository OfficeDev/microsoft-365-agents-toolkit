// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { shared, header, promise, typeChecking } from "../eslint-plugin-teamsfx/config/flat.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));

export default [
  ...shared(__dirname),
  header(),
  promise(__dirname),
  typeChecking(__dirname),
];
