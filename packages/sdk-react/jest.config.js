module.exports = {
  displayName: "client",
  rootDir: "./",
  testEnvironment: "jsdom",
  transform: {
    "\\.tsx?$": [
      "ts-jest",
      {
        tsconfig: "<rootDir>/tsconfig.json",
        diagnostics: {
          ignoreCodes: [151001],
        },
      },
    ],
    "\\.jsx?$": "babel-jest",
  },
  transformIgnorePatterns: [
    "<rootDir>/../sdk/node_modules/(?!@azure/core-auth|@azure/core-http|@microsoft/agents-activity|@microsoft/agents-hosting|@microsoft/agents-hosting-teams|@microsoft/agents-hosting-dialogs)",
  ],
  moduleNameMapper: {
    // Force module uuid to resolve with the CJS entry point, because Jest does not support package.json.exports. See https://github.com/uuidjs/uuid/issues/451
    uuid: require.resolve("uuid"),
    "^react($|/.+)": "<rootDir>/node_modules/react$1",
  },
  preset: "ts-jest",
  setupFiles: ["<rootDir>/jest.setup.js"],
  testMatch: ["<rootDir>/test/*.test.(ts|tsx|js)"],
  collectCoverage: true,
  collectCoverageFrom: ["/src/*.{js,jsx,ts,tsx}", "!<rootDir>/node_modules/"],
  coverageReporters: ["text", "html", "lcov"],
};
