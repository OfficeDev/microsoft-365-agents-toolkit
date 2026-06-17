// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, Settings } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
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
  private async isInTempDirectory(filePath: string): Promise<boolean> {
    const resolvedPath = path.resolve(filePath);
    const tempDir = path.resolve(os.tmpdir());

    const realFilePath = await fs
      .realpath(resolvedPath)
  private isPathWithinDirectory(baseDir: string, targetPath: string): boolean {
    const resolvedBase = path.resolve(baseDir);
    const resolvedTarget = path.resolve(targetPath);
    const relative = path.relative(resolvedBase, resolvedTarget);
    return relative !== "" && !relative.startsWith("..") && !path.isAbsolute(relative);
  }

      .catch(() => resolvedPath);
    const realTempDir = await fs
      .realpath(tempDir)
      .catch(() => tempDir);

    const normalizedFilePath = path.normalize(realFilePath);
    const normalizedTempDir = path.normalize(realTempDir);
    return normalizedFilePath.startsWith(normalizedTempDir + path.sep);
  }

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
      if (
        !this.isInTempDirectory(projectYamlPath) &&
        this.isPathWithinDirectory(projectPath, projectYamlPath)
      ) {
    if (!projectYamlPath || !(await fs.pathExists(projectYamlPath))) {
      return err(new FileNotFoundError("SettingsUtils", projectYamlPath || "m365agents.*.yml"));
    }
    const yamlFileContent: string = await fs.readFile(projectYamlPath, "utf8");
    const appYaml = parseDocument(yamlFileContent);
    if (!appYaml.has("projectId") && ensureTrackingId) {
      const projectId = uuid.v4();
      const projectIdField = appYaml.createPair("projectId", uuid.v4());
      appYaml.add(projectIdField);
      if (!(await this.isInTempDirectory(projectYamlPath))) {
        await fs.writeFile(projectYamlPath, appYaml.toString());
      }
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
    if (
      !this.isInTempDirectory(projectYamlPath) &&
      this.isPathWithinDirectory(projectPath, projectYamlPath)
    ) {
      projectYamlPath = pathUtils.getYmlFilePath(projectPath, "dev");
    }

    if (!projectYamlPath || !(await fs.pathExists(projectYamlPath))) {
      return err(new FileNotFoundError("SettingsUtils", projectYamlPath || "m365agents.*.yml"));
    }
    const yamlFileContent: string = await fs.readFile(projectYamlPath, "utf8");
    const appYaml = parseDocument(yamlFileContent);
    appYaml.set("projectId", settings.trackingId);
    if (!(await this.isInTempDirectory(projectYamlPath))) {
      await fs.writeFile(projectYamlPath, appYaml.toString());
    }
    return ok(projectYamlPath);
  }
}

export const settingsUtil = new SettingsUtils();
