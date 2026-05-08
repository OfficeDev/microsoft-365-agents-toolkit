export default {
  languageOptions: {
    parserOptions: {
      project: ["./tsconfig.eslint.json"],
    },
  },
  rules: {
    "@typescript-eslint/no-unsafe-member-access": "error",
    "@typescript-eslint/no-unsafe-argument": "error",
    "@typescript-eslint/no-unsafe-assignment": "error",
    "@typescript-eslint/no-unsafe-call": "error",
    "@typescript-eslint/no-unsafe-return": "error",
  },
};
