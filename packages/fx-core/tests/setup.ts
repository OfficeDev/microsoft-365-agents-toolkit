// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import "ts-node/register/transpile-only";
import { vi } from "vitest";
import "./component/deps-checker/testGuard";

const generatedArtifactPaths = [
  "path",
  "projectPath",
  "pluginManifestPath",
  path.join("mock", "path", ".kiotabin"),
  path.join("mock", "to", "kiota", ".kiotabin"),
];

async function clearGeneratedTestArtifacts(): Promise<void> {
  await Promise.all(
    generatedArtifactPaths.map(async (relativePath) => {
      const absolutePath = path.resolve(process.cwd(), relativePath);
      await fs.remove(absolutePath);
    })
  );
}

function suppressSourcemapNoise(stream: NodeJS.WriteStream): void {
  const originalWrite = stream.write.bind(stream);
  stream.write = ((chunk: unknown, ...args: unknown[]) => {
    const text =
      typeof chunk === "string" ? chunk : Buffer.isBuffer(chunk) ? chunk.toString("utf8") : "";
    if (text.includes("Sourcemap for ") && text.includes("points to missing source files")) {
      return true;
    }
    return (originalWrite as any)(chunk, ...args);
  }) as typeof stream.write;
}

suppressSourcemapNoise(process.stdout);
suppressSourcemapNoise(process.stderr);

const warningHandler = () => {
  // Suppress runtime warnings from dependencies during UT runs.
};

const mutedConsoleMethods: Array<keyof Console> = ["log", "info", "warn", "error", "debug"];
const restoreConsoleSpies = mutedConsoleMethods.map((method) =>
  vi.spyOn(console, method).mockImplementation(() => undefined)
);

process.on("warning", warningHandler);

beforeAll(async () => {
  // Artifact cleanup is handled by global setup to avoid per-file I/O overhead.
});

afterAll(async () => {
  process.off("warning", warningHandler);
  restoreConsoleSpies.forEach((spy) => spy.mockRestore());
});

// Keep legacy BDD aliases used by existing tests.
Object.assign(globalThis, {
  before: beforeAll,
  after: afterAll,
  context: describe,
});
