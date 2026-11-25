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
}
