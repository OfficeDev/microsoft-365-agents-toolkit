import { defineConfig } from "vitest/config";

export default defineConfig({
  logLevel: "error",
  test: {
    globals: true,
    fileParallelism: true,
    retry: 0,
    include: ["tests/**/*.test.ts"],
    globalSetup: ["tests/globalSetup.ts"],
    setupFiles: ["tests/setup.ts"],
    environment: "node",
    testTimeout: 60000,
    hookTimeout: 60000,
    teardownTimeout: 60000,
    onConsoleLog(log) {
      if (log.includes("Sourcemap for ") && log.includes("points to missing source files")) {
        return false;
      }
    },
    coverage: {
      provider: "istanbul",
      reportsDirectory: "coverage",
      reporter: ["html", "text", "json-summary", "cobertura", "lcov"],
      all: true,
      include: ["src/**/*.ts"],
      exclude: ["src/component/deps-checker/**/*", "src/question/generator.ts"],
      thresholds: {
        lines: 72.7,
      },
    },
  },
});
