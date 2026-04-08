import shared from "../eslint-plugin-teamsfx/config/shared.mjs";
import header from "../eslint-plugin-teamsfx/config/header.mjs";
import promise from "../eslint-plugin-teamsfx/config/promise.mjs";
import typeConfig from "../eslint-plugin-teamsfx/config/type.mjs";

export default [
  {
    ignores: ["src/chat/api/**"],
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
  { files: ["src/**/*.ts"], ...promise },
  { files: ["src/**/*.ts"], ...typeConfig },
];
