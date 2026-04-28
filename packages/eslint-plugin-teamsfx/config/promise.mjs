export default {
  languageOptions: {
    parserOptions: {
      project: ["./tsconfig.eslint.json"],
    },
  },
  rules: {
    "@typescript-eslint/await-thenable": "error",
    "@typescript-eslint/no-floating-promises": "warn",
    "@typescript-eslint/no-misused-promises": [
      "error",
      { checksVoidReturn: { arguments: false } },
    ],
    "@typescript-eslint/require-await": "error",
  },
};
