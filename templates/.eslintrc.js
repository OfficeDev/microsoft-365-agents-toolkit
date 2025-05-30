module.exports = {
  extends: ["../packages/eslint-plugin-teamsfx/config/shared.js"],
  overrides: [
    {
      files: ["vsc/**/*.ts"],
      rules: {
        "no-redeclare": "off", // Allow redeclaration in VS Code extension files
      },
    },
  ],
};
