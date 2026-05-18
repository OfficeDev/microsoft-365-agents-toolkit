// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";

export interface ConvertOpenPluginInputs extends Inputs {
  /** Path to the Open Plugin directory. */
  path: string;
  /** Destination project folder. */
  output?: string;
  /** developer.privacyUrl for the generated manifest. */
  "privacy-url": string;
  /** developer.termsOfUseUrl for the generated manifest. */
  "terms-url": string;
  /** developer.websiteUrl override. */
  "website-url"?: string;
  /** Override the deterministic UUIDv5 generated for the manifest id. */
  "app-id"?: string;
  /** Default auth type for discovered MCP servers. */
  "default-auth-type"?: "Auto" | "None" | "OAuthPluginVault" | "ApiKeyPluginVault";
  /** Full reverse-DNS package name. */
  "package-name"?: string;
}
