import shared from "../eslint-plugin-teamsfx/config/shared.mjs";
import header from "../eslint-plugin-teamsfx/config/header.mjs";

export default [
  {
    ignores: ["office-xml-addin/**", "word-xml-addin/**", "resource/**"],
  },
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
