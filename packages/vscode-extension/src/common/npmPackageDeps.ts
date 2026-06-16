// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Centralized dependency injection for read-only npm packages.
 *
 * This module wraps all external npm package methods that need to be mocked in tests.
 * Internal modules (fileSystemUtils, appDefinitionUtils, etc.) should be imported directly.
 *
 * Principle:
 * - ✅ npm packages (fs-extra, @microsoft/teamsfx-core, etc.) → goes here
 * - ✅ internal modules → import directly, mock with vi.mock() in tests
 */

import fs from "fs-extra";
import { dotenvUtil } from "@microsoft/teamsfx-core/build/component/utils/envUtil";

/**
 * File system operations adapter for fs-extra
 */
export const fsAdapter = {
  pathExists: (filePath: string) => fs.pathExists(filePath),
  readFileSync: (filePath: string, encoding: BufferEncoding) => fs.readFileSync(filePath, encoding),
  writeFileSync: (filePath: string, content: string, encoding?: BufferEncoding) =>
    fs.writeFileSync(filePath, content, encoding || "utf-8"),
  existsSync: (filePath: string) => fs.existsSync(filePath),
  mkdirSync: (dirPath: string, options?: any) => fs.mkdirSync(dirPath, options),
  closeSync: (fd: number) => fs.closeSync(fd),
};

/**
 * Environment file parsing adapter for @microsoft/teamsfx-core
 */
export const envParseAdapter = {
  deserializeDotenv: (content: string) => dotenvUtil.deserialize(content),
};
