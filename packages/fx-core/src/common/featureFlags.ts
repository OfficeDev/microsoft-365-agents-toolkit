// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Determine whether feature flag is enabled based on environment variable setting
export function isFeatureFlagEnabled(featureFlagName: string, defaultValue = false): boolean {
  const flag = process.env[featureFlagName];
  if (flag === undefined) {
    return defaultValue; // allows consumer to set a default value when environment variable not set
  } else {
    return flag === "1" || flag.toLowerCase() === "true"; // can enable feature flag by set environment variable value to "1" or "true"
  }
}
export class FeatureFlagName {
  static readonly CLIDotNet = "TEAMSFX_CLI_DOTNET";
  static readonly OfficeMetaOS = "TEAMSFX_OFFICE_METAOS";
  static readonly SampleConfigBranch = "TEAMSFX_SAMPLE_CONFIG_BRANCH";
  static readonly TestTool = "TEAMSFX_TEST_TOOL";
  static readonly METestTool = "TEAMSFX_ME_TEST_TOOL";
  static readonly TeamsFxRebranding = "TEAMSFX_REBRANDING";
  static readonly TdpTemplateCliTest = "TEAMSFX_TDP_TEMPLATE_CLI_TEST";
  static readonly AsyncAppValidation = "TEAMSFX_ASYNC_APP_VALIDATION";
  static readonly NewProjectType = "TEAMSFX_NEW_PROJECT_TYPE";
  static readonly ChatParticipant = "TEAMSFX_CHAT_PARTICIPANT";
  static readonly ChatParticipantUIEntries = "TEAMSFX_CHAT_PARTICIPANT_ENTRIES";
  static readonly SMEOAuth = "SME_OAUTH";
  static readonly ShowDiagnostics = "TEAMSFX_SHOW_DIAGNOSTICS";
  static readonly TelemetryTest = "TEAMSFX_TELEMETRY_TEST";
  static readonly DevTunnelTest = "TEAMSFX_DEV_TUNNEL_TEST";
  static readonly SyncManifest = "TEAMSFX_SYNC_MANIFEST";
  static readonly KiotaIntegration = "TEAMSFX_KIOTA_INTEGRATION";
  static readonly CEAEnabled = "TEAMSFX_CEA_ENABLED";
  static readonly BuilderAPIEnabled = "TEAMSFX_BUILDER_API";
}

export interface FeatureFlag {
  name: string;
  defaultValue: string;
  description?: string;
}

export class FeatureFlags {
  static readonly CLIDotNet = { name: FeatureFlagName.CLIDotNet, defaultValue: "false" };
  static readonly TestTool = { name: FeatureFlagName.TestTool, defaultValue: "true" };
  static readonly METestTool = { name: FeatureFlagName.METestTool, defaultValue: "true" };
  static readonly OfficeMetaOS = {
    name: FeatureFlagName.OfficeMetaOS,
    defaultValue: "false",
  };
  static readonly TdpTemplateCliTest = {
    name: FeatureFlagName.TdpTemplateCliTest,
    defaultValue: "false",
  };
  static readonly AsyncAppValidation = {
    name: FeatureFlagName.AsyncAppValidation,
    defaultValue: "true",
  };
  static readonly NewProjectType = { name: FeatureFlagName.NewProjectType, defaultValue: "true" };
  static readonly ChatParticipant = {
    name: FeatureFlagName.ChatParticipant,
    defaultValue: "false",
  };
  static readonly ChatParticipantUIEntries = {
    name: FeatureFlagName.ChatParticipantUIEntries,
    defaultValue: "false",
  };
  static readonly SMEOAuth = { name: FeatureFlagName.SMEOAuth, defaultValue: "false" };
  static readonly ShowDiagnostics = {
    name: FeatureFlagName.ShowDiagnostics,
    defaultValue: "false",
  };
  static readonly TelemetryTest = {
    name: FeatureFlagName.TelemetryTest,
    defaultValue: "false",
  };
  static readonly DevTunnelTest = {
    name: FeatureFlagName.DevTunnelTest,
    defaultValue: "false",
  };
  static readonly SyncManifest = {
    name: FeatureFlagName.SyncManifest,
    defaultValue: "false",
  };
  static readonly KiotaIntegration = {
    name: FeatureFlagName.KiotaIntegration,
    defaultValue: "false",
  };
  static readonly CEAEnabled = {
    name: FeatureFlagName.CEAEnabled,
    defaultValue: "false",
  };
  static readonly BuilderAPIEnabled = {
    name: FeatureFlagName.BuilderAPIEnabled,
    defaultValue: "false",
  };
}

export class FeatureFlagManager {
  getBooleanValue(featureFlag: FeatureFlag): boolean {
    return isFeatureFlagEnabled(
      featureFlag.name,
      featureFlag.defaultValue === "true" || featureFlag.defaultValue === "1"
    );
  }
  setBooleanValue(featureFlag: FeatureFlag, value: boolean): void {
    process.env[featureFlag.name] = value ? "true" : "false";
  }
  getStringValue(featureFlag: FeatureFlag): string {
    return process.env[featureFlag.name] || featureFlag.defaultValue;
  }
  list(): FeatureFlag[] {
    return Object.values(FeatureFlags);
  }
  listEnabled(): string[] {
    return this.list()
      .filter((f) => isFeatureFlagEnabled(f.name))
      .map((f) => f.name);
  }
}

export const featureFlagManager = new FeatureFlagManager();
