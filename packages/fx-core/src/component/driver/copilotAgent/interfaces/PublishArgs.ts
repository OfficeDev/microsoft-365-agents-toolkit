// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CopilotAgentPublishArgs {
  /**
   * Zipped app package path
   */
  appPackagePath: string;

  /**
   * Publishing scope. Defaults to "Personal".
   * - Personal: Publish to personal scope
   * - Shared: Publish to shared scope
   * - Tenant: Publish to tenant app catalog
   */
  scope?: string;

  /**
   * Prepared publication journal path written by `spfx/preparePublication`.
   * The prepared mode is selected when any of the prepared inputs is set; all of them are then required.
   */
  preparedPublicationPath?: string;

  /**
   * Prepared publication ID written by `spfx/preparePublication`.
   */
  publicationId?: string;

  /**
   * SHA-256 of the source SharePoint package, written by `spfx/preparePublication`.
   */
  sourcePackageSha256?: string;

  /**
   * SHA-256 of the source app package, written by `spfx/preparePublication`.
   */
  sourceAppPackageSha256?: string;

  /**
   * SHA-256 of the extracted API plugin, written by `spfx/preparePublication`.
   */
  pluginSha256?: string;

  /**
   * SHA-256 of the prepared SharePoint deployment package, written by `spfx/preparePublication`.
   */
  sharePointPackageSha256?: string;

  /**
   * SHA-256 of the prepared app package, written by `spfx/preparePublication`.
   */
  appPackageSha256?: string;
}
