module.exports = {
  parser: "@typescript-eslint/parser",
  parserOptions: {
    tsconfigRootDir: __dirname,
  },
  extends: ["../eslint-plugin-teamsfx/config/shared.js"],
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
