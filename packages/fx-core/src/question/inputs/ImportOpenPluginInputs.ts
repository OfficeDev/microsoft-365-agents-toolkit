// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";

export interface ImportOpenPluginInputs extends Inputs {
  /** Path to the Open Plugin directory. */
  path: string;
  /** Destination project folder. */
  output?: string;
  /**
   * developer.privacyUrl for the generated manifest. Optional when the source
   * plugin.json carries an `x-microsoft-365-agents-toolkit` extension block
   * with a previously-exported developer.privacyUrl.
   */
  "privacy-url"?: string;
  /**
   * developer.termsOfUseUrl for the generated manifest. Optional when the
   * source plugin.json carries an `x-microsoft-365-agents-toolkit` extension
   * block with a previously-exported developer.termsOfUseUrl.
   */
  "terms-url"?: string;
  /** developer.websiteUrl override. */
  "website-url"?: string;
  /** Override the deterministic UUIDv5 generated for the manifest id. */
  "app-id"?: string;
  /** Default auth type for discovered MCP servers. */
  "default-auth-type"?: "Auto" | "None" | "OAuthPluginVault" | "ApiKeyPluginVault";
  /** Full reverse-DNS package name. */
  "package-name"?: string;
}
