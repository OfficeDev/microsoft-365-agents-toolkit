import tseslint from "typescript-eslint";
import eslintPluginPrettierRecommended from "eslint-plugin-prettier/recommended";
import importPlugin from "eslint-plugin-import-x";
import noSecrets from "eslint-plugin-no-secrets";
import globals from "globals";

export default [
  {
    ignores: ["build/**", "lib/**", "dist/**", "out/**", "coverage/**"],
  },
  {
    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.es2021,
        ...globals.node,
      },
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: "module",
      },
    },
  },
  ...tseslint.configs.recommended,
  eslintPluginPrettierRecommended,
  importPlugin.flatConfigs.recommended,
  {
    ...importPlugin.flatConfigs.typescript,
    settings: {
      ...importPlugin.flatConfigs.typescript.settings,
      "import-x/resolver": { node: true },
    },
  },
  {
    plugins: {
      "no-secrets": noSecrets,
    },
    rules: {
      quotes: [
        "error",
        "double",
        { allowTemplateLiterals: true, avoidEscape: true },
      ],
      semi: ["error", "always"],
      "@typescript-eslint/no-require-imports": "off",
      "@typescript-eslint/no-empty-function": "off",
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/no-unused-vars": "warn",
      "@typescript-eslint/no-unused-expressions": "off",
      "@typescript-eslint/no-duplicate-enum-values": "warn",
      "@typescript-eslint/no-unsafe-declaration-merging": "warn",
      "import-x/no-cycle": [
        "error",
        {
          maxDepth: Infinity,
          ignoreExternal: true,
        },
      ],
      "import-x/no-unresolved": ["warn"],
      "no-secrets/no-secrets": [
        "warn",
        {
          additionalRegexes: {
            "Basic Auth": "Authorization: Basic [A-Za-z0-9+/=]*",
            "Common Pattern":
              "^(?=.*[A-Za-z])(?=.*[0-9])(?=.*[@$!%*#?&])[A-Za-z0-9@$!%*#?&~-]{8,}$",
          },
        },
      ],
    },
  },
];
