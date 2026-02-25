"use strict";

module.exports = {
  env: {
    browser: true,
    es6: true,
    node: true,
  },
  parserOptions: {
    ecmaVersion: 2018,
    sourceType: "module",
  },
  extends: [
    "plugin:@typescript-eslint/recommended",
    "prettier",
    "plugin:import/errors",
    "plugin:import/warnings",
    "plugin:import/typescript",
  ],
  plugins: [
    "@typescript-eslint/eslint-plugin",
    "prettier",
    "no-secrets",
    "header",
  ],
  rules: {
    "prettier/prettier": "error",
    quotes: [
      "error",
      "double",
      { allowTemplateLiterals: true, avoidEscape: true },
    ],
    semi: ["error", "always"],
    "@typescript-eslint/no-require-imports": 0,
    "@typescript-eslint/no-empty-function": 0,
    "import/no-cycle": [
      "error",
      {
        maxDepth: Infinity,
        ignoreExternal: true,
      },
    ],
    "import/no-unresolved": ["warn"],
    "no-secrets/no-secrets": [
      "warn",
      {
        tolerance: 4.5,
        additionalRegexes: {
          "Basic Auth": "Authorization: Basic [A-Za-z0-9+/=]*",
          "Common Pattern":
            "^(?=.*[A-Za-z])(?=.*[0-9])(?=.*[@$!%*#?&])[A-Za-z0-9@$!%*#?&~-]{8,}$",
        },
        ignoreContent: [
          "^[A-Z][a-zA-Z]*Wrapper$",
          "^[A-Z][a-zA-Z]*Manifest$"
        ],
      },
    ],
  },
};
