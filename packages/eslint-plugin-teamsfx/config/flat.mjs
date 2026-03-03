// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import tseslint from "typescript-eslint";
import eslintConfigPrettier from "eslint-config-prettier";
import importPlugin from "eslint-plugin-import";
import noSecretsPlugin from "eslint-plugin-no-secrets";
import headerPlugin from "eslint-plugin-header";
import globals from "globals";
import { fixupPluginRules } from "@eslint/compat";

/**
 * Patch plugin rules to provide proper schemas where plugins have schema: false.
 * ESLint 9 strictly validates rule options against schemas.
 */
function patchRuleSchema(plugin, patches) {
  const patched = { ...plugin, rules: { ...plugin.rules } };
  for (const [ruleName, schema] of Object.entries(patches)) {
    if (patched.rules[ruleName]) {
      const rule = patched.rules[ruleName];
      patched.rules[ruleName] = {
        ...rule,
        meta: { ...rule.meta, schema },
      };
    }
  }
  return patched;
}

const patchedHeaderPlugin = patchRuleSchema(headerPlugin, {
  header: [
    { type: "string", enum: ["line", "block"] },
    {
      oneOf: [
        { type: "string" },
        { type: "array", items: { type: "string" } },
      ],
    },
  ],
});

const patchedNoSecretsPlugin = patchRuleSchema(noSecretsPlugin, {
  "no-secrets": [
    {
      type: "object",
      properties: {
        tolerance: { type: "number" },
        additionalRegexes: { type: "object" },
        ignoreContent: { type: "array", items: { type: "string" } },
        additionalDelimiters: { type: "array", items: { type: "string" } },
      },
      additionalProperties: false,
    },
  ],
});

const plugins = {
  "@typescript-eslint": tseslint.plugin,
  "no-secrets": fixupPluginRules(patchedNoSecretsPlugin),
  header: fixupPluginRules(patchedHeaderPlugin),
  import: fixupPluginRules(importPlugin),
};

/**
 * Shared base configuration.
 * @param {string} tsconfigRootDir
 * @returns {import("eslint").Linter.Config[]}
 */
export function shared(tsconfigRootDir) {
  return [
    {
      languageOptions: {
        parser: tseslint.parser,
        globals: {
          ...globals.browser,
          ...globals.es2017,
          ...globals.node,
        },
        parserOptions: {
          ecmaVersion: 2018,
          sourceType: "module",
          tsconfigRootDir,
        },
      },
      plugins,
      settings: {
        "import/parsers": {
          "@typescript-eslint/parser": [".ts", ".tsx", ".d.ts"],
        },
        "import/resolver": {
          node: {
            extensions: [".ts", ".tsx", ".d.ts", ".js", ".jsx"],
          },
        },
      },
    },
    ...tseslint.configs.recommended,
    eslintConfigPrettier,
    {
      rules: {
        quotes: [
          "error",
          "double",
          { allowTemplateLiterals: true, avoidEscape: true },
        ],
        semi: ["error", "always"],
        "@typescript-eslint/no-require-imports": "off",
        "@typescript-eslint/no-empty-function": "off",
        // Restore v7 recommended severities (v8 changed these from "warn" to "error")
        "@typescript-eslint/no-unused-vars": [
          "warn",
          { args: "none", ignoreRestSiblings: true },
        ],
        "@typescript-eslint/no-explicit-any": "warn",
        // Not in v7 recommended; commonly triggered by chai assertions
        "@typescript-eslint/no-unused-expressions": "off",
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
              "^[A-Z][a-zA-Z]*Manifest$",
            ],
          },
        ],
      },
    },
  ];
}

/**
 * Header comment enforcement for source files.
 * @param {string[]} files
 * @returns {import("eslint").Linter.Config}
 */
export function header(files = ["src/**/*.ts"]) {
  return {
    files,
    rules: {
      "header/header": [
        "error",
        "line",
        [
          " Copyright (c) Microsoft Corporation.",
          " Licensed under the MIT license.",
        ],
      ],
    },
  };
}

/**
 * Promise/async rules requiring type checking.
 * @param {string} tsconfigRootDir
 * @param {string[]} files
 * @returns {import("eslint").Linter.Config}
 */
export function promise(tsconfigRootDir, files = ["src/**/*.ts"]) {
  return {
    files,
    languageOptions: {
      parserOptions: {
        project: ["./tsconfig.eslint.json"],
        tsconfigRootDir,
      },
    },
    rules: {
      "@typescript-eslint/await-thenable": "error",
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      "@typescript-eslint/require-await": "error",
    },
  };
}

/**
 * Type checking rules.
 * @param {string} tsconfigRootDir
 * @param {string[]} files
 * @returns {import("eslint").Linter.Config}
 */
export function typeChecking(tsconfigRootDir, files = ["src/**/*.ts"]) {
  return {
    files,
    languageOptions: {
      parserOptions: {
        project: ["./tsconfig.eslint.json"],
        tsconfigRootDir,
      },
    },
    rules: {
      "@typescript-eslint/no-unnecessary-type-assertion": "error",
      "@typescript-eslint/no-for-in-array": "error",
      "@typescript-eslint/no-implied-eval": "error",
      "@typescript-eslint/restrict-plus-operands": "error",
      "@typescript-eslint/restrict-template-expressions": "error",
    },
  };
}
