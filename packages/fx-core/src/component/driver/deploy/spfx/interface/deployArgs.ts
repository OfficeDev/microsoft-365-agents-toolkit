// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface DeploySPFxArgs {
  /**
   * create app catalog if there is no valid one in the
   * current Microsoft 365 tenant, default to 'false'
   */
  createAppCatalogIfNotExist: boolean;

  /**
   * the 'package-solution.json' path of the SPFx project,
   * will use it to get the bundled zipped file path.
   */
  packageSolutionPath: string;

  /**
   * Prepared publication journal path. When supplied with publicationId,
   * deploys only the immutable package recorded in that journal.
   */
  preparedPublicationPath?: string;

  /**
   * Prepared publication identifier emitted by spfx/preparePublication.
   */
  publicationId?: string;

  preparedPackagePath?: string;
  sourcePackageSha256?: string;
  sourceAppPackageSha256?: string;
  pluginSha256?: string;
  sharePointPackageSha256?: string;
  appPackageSha256?: string;
}
