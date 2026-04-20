// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CreateAppPackageArgs {
  /**
   * Teams app manifest template path
   */
  manifestPath: string;

  /**
   * Zipped app package path
   */
  outputZipPath: string;

  /**
   * Manifest file path. This parameter is used when teamspp yaml version <= 1.6
   */
  outputJsonPath?: string;

  /**
   * Folder path where output files should be put.  This parameter is used when teamspp yaml version > 1.6
   */
  outputFolder?: string;

  /**
   * Maximum allowed size (in bytes) for the generated zip package.
   * If set, the action will fail when the zip exceeds this limit.
   * Useful for enforcing MOS/Titles API upload limits (e.g. 10 MB = 10485760).
   */
  maxPackageSizeInBytes?: number;
}
