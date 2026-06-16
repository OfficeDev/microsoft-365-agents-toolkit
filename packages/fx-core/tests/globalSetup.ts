// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";

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

export default async function globalSetup() {
  await clearGeneratedTestArtifacts();

  return async () => {
    await clearGeneratedTestArtifacts();
  };
}
