// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import js from "@eslint/js";
import globals from "globals";

export default [
  js.configs.recommended,
  {
    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.es2017,
        ...globals.node,
      },
    },
  },
  {
    files: ["tests/**/*.js"],
    languageOptions: {
      globals: {
        ...globals.mocha,
      },
    },
  },
];
