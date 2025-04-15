module.exports = {
  parser: "@typescript-eslint/parser",
  parserOptions: {
    tsconfigRootDir: __dirname,
  },
  extends: ["../eslint-plugin-teamsfx/config/shared.js"],
  ignorePatterns: ["src/generated-types"], // Add folders to ignore here
  overrides: [
    {
      files: ["src/**/*.ts"],
      extends: [
        "../eslint-plugin-teamsfx/config/header.js",
        "../eslint-plugin-teamsfx/config/promise.js",
        "../eslint-plugin-teamsfx/config/type.js",
      ],
    },
  ],
};
