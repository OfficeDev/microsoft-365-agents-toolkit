// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CreateDcrArgs {
  name: string; // The display name of the DCR config; becomes clientName; max 128 chars
  appId?: string; // Teams app id; required only when applicableToApps is "SpecificApp"
  mcpResourceUrl?: string;
  wellKnownAuthorizationServer?: string;
  resource?: string;
  targetUrlsShouldStartWith: string[];
  applicableToApps?: string; // Which apps can use this config. Values: "SpecificApp" | "AnyApp". Default: "AnyApp".
  targetAudience?: string; // Which tenants can use this config. Values: "HomeTenant" | "AnyTenant". Default: "HomeTenant".
}
