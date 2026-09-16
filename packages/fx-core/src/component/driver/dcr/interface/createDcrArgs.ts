// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CreateDcrArgs {
  name: string; // The display name of the DCR config; becomes clientName; max 128 chars
  appId?: string; // Published Microsoft 365 app ID (MOSAppId), not Teams manifest ID; required for "SpecificApp".
  wellKnownAuthorizationServer: string; // URL to the AS's RFC 8414 metadata document; required
  targetUrlsShouldStartWith?: string[]; // URL prefixes allowed as outbound destinations
  applicableToApps?: string; // Which apps can use this config. Values: "SpecificApp" | "AnyApp". Default: "AnyApp".
  targetAudience?: string; // Which tenants can use this config. Values: "HomeTenant" | "AnyTenant". Default: "HomeTenant".
}
