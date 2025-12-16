// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, Settings } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as uuid from "uuid";
import { parseDocument } from "yaml";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";
import { globalVars } from "../../common/globalVars";
import {
  Component,
  sendTelemetryEvent,
  TelemetryEvent,
  TelemetryProperty,
} from "../../common/telemetry";
import { FileNotFoundError } from "../../error/common";
import { pathUtils } from "./pathUtils";

class SettingsUtils {
  async readSettings(
    projectPath: string,
    ensureTrackingId = true
  ): Promise<Result<Settings, FxError>> {
    let projectYamlPath: string | undefined;
    if (featureFlagManager.getBooleanValue(FeatureFlags.GenerateConfigFiles)) {
      projectYamlPath = pathUtils.getAvailableYmlFilePath(projectPath);
    } else {
      projectYamlPath = pathUtils.getYmlFilePath(projectPath, "dev");
    }

    if (!projectYamlPath || !(await fs.pathExists(projectYamlPath))) {
      return err(new FileNotFoundError("SettingsUtils", projectYamlPath || "m365agents.*.yml"));
    }
    const yamlFileContent: string = await fs.readFile(projectYamlPath, "utf8");
    const appYaml = parseDocument(yamlFileContent);
    if (!appYaml.has("projectId") && ensureTrackingId) {
      const projectId = uuid.v4();
      const projectIdField = appYaml.createPair("projectId", uuid.v4());
      appYaml.add(projectIdField);
      await fs.writeFile(projectYamlPath, appYaml.toString()); // only write yaml file once instead of write yaml file after every command
      sendTelemetryEvent(Component.core, TelemetryEvent.FillProjectId, {
        [TelemetryProperty.ProjectId]: projectId,
      });
    }
    const projectSettings: Settings = {
      trackingId: appYaml.get("projectId") as string,
      version: appYaml.get("version") as string,
    };

    globalVars.trackingId = projectSettings.trackingId; // set trackingId to globalVars
    return ok(projectSettings);
  }
  async writeSettings(projectPath: string, settings: Settings): Promise<Result<string, FxError>> {
    let projectYamlPath: string | undefined;
    if (featureFlagManager.getBooleanValue(FeatureFlags.GenerateConfigFiles)) {
      projectYamlPath = pathUtils.getAvailableYmlFilePath(projectPath);
    } else {
      projectYamlPath = pathUtils.getYmlFilePath(projectPath, "dev");
    }

    if (!projectYamlPath || !(await fs.pathExists(projectYamlPath))) {
      return err(new FileNotFoundError("SettingsUtils", projectYamlPath || "m365agents.*.yml"));
    }
    const yamlFileContent: string = await fs.readFile(projectYamlPath, "utf8");
    const appYaml = parseDocument(yamlFileContent);
    appYaml.set("projectId", settings.trackingId);
    await fs.writeFile(projectYamlPath, appYaml.toString());
    return ok(projectYamlPath);
  }
}

export const settingsUtil = new SettingsUtils();
