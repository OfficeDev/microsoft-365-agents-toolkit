// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface PublishAppPackageArgs {
  /**
   * Zipped app package path
   */
  appPackagePath: string;

  /**
   * Prepared publication journal path emitted by spfx/preparePublication.
   */
  preparedPublicationPath?: string;

  /**
   * Prepared publication identifier emitted by spfx/preparePublication.
   */
  publicationId?: string;

  sourcePackageSha256?: string;
  sourceAppPackageSha256?: string;
  pluginSha256?: string;
  sharePointPackageSha256?: string;
  appPackageSha256?: string;
}
