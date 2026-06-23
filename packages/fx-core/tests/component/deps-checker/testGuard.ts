// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import { afterEach, beforeEach, expect } from "vitest";
import * as fetchHelper from "../../../src/common/fetchHelper";
import { cpUtils } from "../../../src/component/deps-checker/util/cpUtils";
import * as downloadHelper from "../../../src/component/deps-checker/util/downloadHelper";

const originalExecuteCommand = cpUtils.executeCommand;
const originalFetch = fetchHelper.default;
const originalDownloadToTempFile = downloadHelper.downloadToTempFile;
const originalUnzip = downloadHelper.unzip;

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

  fetchHelper.default = (async () => {
    throw createUnmockedDependencyError("fetchHelper.default");
  }) as typeof fetchHelper.default;

  downloadHelper.downloadToTempFile = (async () => {
    throw createUnmockedDependencyError("downloadHelper.downloadToTempFile");
  }) as typeof downloadHelper.downloadToTempFile;

  downloadHelper.unzip = (async () => {
    throw createUnmockedDependencyError("downloadHelper.unzip");
  }) as typeof downloadHelper.unzip;
});

afterEach(() => {
  if (!isDepsCheckerTestFile()) {
    return;
  }

  cpUtils.executeCommand = originalExecuteCommand;
  fetchHelper.default = originalFetch;
  downloadHelper.downloadToTempFile = originalDownloadToTempFile;
  downloadHelper.unzip = originalUnzip;
});
