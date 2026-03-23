import shared from "../eslint-plugin-teamsfx/config/shared.mjs";
import header from "../eslint-plugin-teamsfx/config/header.mjs";

export default [
  ...shared,
  {
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },
  { files: ["src/**/*.ts"], ...header },
];
