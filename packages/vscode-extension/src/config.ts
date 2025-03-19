// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as vscode from "vscode";
import { CONFIGURATION_PREFIX, ConfigurationKey, EnableMicrosoftKiota } from "./constants";
import VsCodeLogInstance from "./commonlib/log";
import { LogLevel } from "@microsoft/teamsfx-api";
import { ExtTelemetry } from "./telemetry/extTelemetry";
import { TelemetryEvent } from "./telemetry/extTelemetryEvents";
import { FeatureFlags } from "@microsoft/teamsfx-core";
import { validateKiotaInstallation } from "./handlers/lifecycleHandlers";
import { VS_CODE_UI } from "./qm/vsc_ui";
import { localize } from "./utils/localizeUtils";

export class ConfigManager {
  registerConfigChangeCallback() {
    this.loadConfigs();
    vscode.workspace.onDidChangeConfiguration?.(this.changeConfigCallback.bind(this));
  }
  loadConfigs() {
    this.loadLogLevel();
    this.loadFeatureFlags();
    const vscConfigs: { [p: string]: string } = {};
    Object.values(ConfigurationKey).forEach((value) => {
      vscConfigs[value] = this.getConfiguration(value, "").toString();
    });
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Configuration, {
      ...vscConfigs,
    });
  }
  loadFeatureFlags() {
    process.env["TEAMSFX_BICEP_ENV_CHECKER_ENABLE"] = this.getConfiguration(
      ConfigurationKey.BicepEnvCheckerEnable,
      false
    ).toString();
    process.env[FeatureFlags.KiotaIntegration.name] = (
      this.getConfiguration(
        ConfigurationKey.EnableMicrosoftKiotaString,
        EnableMicrosoftKiota.undefined
      ).toString() === EnableMicrosoftKiota.enabled
    ).toString();
    process.env[FeatureFlags.CEAEnabled.name] = this.getConfiguration(
      ConfigurationKey.EnableCEA,
      false
    ).toString();
  }
  loadLogLevel() {
    const logLevel = this.getConfiguration(ConfigurationKey.LogLevel, "Info") as string;
    if (logLevel === "Debug") {
      VsCodeLogInstance.logLevel = LogLevel.Debug;
    } else if (logLevel === "Verbose") {
      VsCodeLogInstance.logLevel = LogLevel.Verbose;
    } else {
      VsCodeLogInstance.logLevel = LogLevel.Info;
    }
  }
  getConfiguration(key: string, defaultValue: boolean | string): boolean | string {
    const configuration: vscode.WorkspaceConfiguration =
      vscode.workspace.getConfiguration(CONFIGURATION_PREFIX);
    return configuration.get<boolean | string>(key, defaultValue);
  }
  changeConfigCallback(event: vscode.ConfigurationChangeEvent) {
    if (event.affectsConfiguration(CONFIGURATION_PREFIX)) {
      this.loadConfigs();
    }
  }
  async checkKiotaInstallation() {
    const configuration: vscode.WorkspaceConfiguration =
      vscode.workspace.getConfiguration(CONFIGURATION_PREFIX);
    const currentConfig = configuration.get(ConfigurationKey.EnableMicrosoftKiotaString);
    if (currentConfig === EnableMicrosoftKiota.undefined && validateKiotaInstallation()) {
      const previousConfig = configuration.get(ConfigurationKey.EnableMicrosoftKiota);
      if (previousConfig) {
        await configuration.update(
          ConfigurationKey.EnableMicrosoftKiotaString,
          EnableMicrosoftKiota.enabled
        );
      } else {
        // pop up question to ask if user want to enable kiota
        const res = await VS_CODE_UI.showMessage(
          "warn",
          localize("teamstoolkit.config.enableKiota"),
          true,
          localize("teamstoolkit.config.enableKiota.yes"),
          localize("teamstoolkit.config.enableKiota.no")
        );
        await configuration.update(
          ConfigurationKey.EnableMicrosoftKiotaString,
          res.isOk() && res.value === "Yes"
            ? EnableMicrosoftKiota.enabled
            : EnableMicrosoftKiota.disabled
        );
      }
      this.loadFeatureFlags();
    }
  }
}

export const configMgr = new ConfigManager();
