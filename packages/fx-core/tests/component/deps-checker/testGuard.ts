// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import { afterEach, beforeEach, expect } from "vitest";
import { testToolCheckerDeps } from "../../../src/component/deps-checker/internal/testToolChecker";
import { cpUtils } from "../../../src/component/deps-checker/util/cpUtils";

const originalExecuteCommand = cpUtils.executeCommand;
const originalFetch = testToolCheckerDeps.fetch;
const originalDownloadToTempFile = testToolCheckerDeps.downloadToTempFile;
const originalUnzip = testToolCheckerDeps.unzip;

function isDepsCheckerTestFile(): boolean {
  const testPath = expect.getState().testPath ?? "";
  const marker = `${path.sep}tests${path.sep}component${path.sep}deps-checker${path.sep}`;
  return testPath.includes(marker);
}

function createUnmockedDependencyError(name: string): Error {
  return new Error(
    `Unmocked dependency call detected: ${name}. ` +
      "Please mock this dependency in the test to keep UT deterministic and fast."
  );
}

beforeEach(() => {
  if (!isDepsCheckerTestFile()) {
    return;
  }

  cpUtils.executeCommand = (async () => {
    throw createUnmockedDependencyError("cpUtils.executeCommand");
  }) as typeof cpUtils.executeCommand;

  testToolCheckerDeps.fetch = (async () => {
    throw createUnmockedDependencyError("testToolCheckerDeps.fetch");
  }) as typeof testToolCheckerDeps.fetch;

  testToolCheckerDeps.downloadToTempFile = (async () => {
    throw createUnmockedDependencyError("testToolCheckerDeps.downloadToTempFile");
  }) as typeof testToolCheckerDeps.downloadToTempFile;

  testToolCheckerDeps.unzip = (async () => {
    throw createUnmockedDependencyError("testToolCheckerDeps.unzip");
  }) as typeof testToolCheckerDeps.unzip;
});

afterEach(() => {
  if (!isDepsCheckerTestFile()) {
    return;
  }

  cpUtils.executeCommand = originalExecuteCommand;
  testToolCheckerDeps.fetch = originalFetch;
  testToolCheckerDeps.downloadToTempFile = originalDownloadToTempFile;
  testToolCheckerDeps.unzip = originalUnzip;
});
