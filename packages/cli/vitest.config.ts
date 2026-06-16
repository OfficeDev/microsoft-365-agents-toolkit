import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    globals: true,
    fileParallelism: true,
    include: ["tests/unit/**/*.tests.ts"],
    setupFiles: ["tests/unit/setup.ts"],
    environment: "node",
    testTimeout: 30000,
    hookTimeout: 30000,
    teardownTimeout: 30000,
    coverage: {
      provider: "istanbul",
      reportsDirectory: "coverage",
      reporter: ["html", "text", "json-summary", "cobertura", "lcov"],
      all: true,
      include: ["src/**/*.ts", "src/**/*.js"],
      exclude: [
        "src/cmds/preview/depsChecker/**/*",
        "src/cmds/preview/preview.ts",
        "src/cmds/preview/npmInstallHandler.ts",
        "src/commonlib/*Login.ts",
        "src/commonlib/*LoginCI.ts",
        "src/commonlib/*LoginUserPassword.ts",
        "src/index.ts",
        "cli.js",
        ".eslintrc.js",
        ".prettierrc.js",
        ".mocharc.js",
      ],
      thresholds: {
        lines: 80,
      },
    },
  },
});
