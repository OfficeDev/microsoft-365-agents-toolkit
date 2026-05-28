// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as vscode from "vscode";
import * as os from "os";
import path from "path";
// eslint-disable-next-line import/default
import Reporter from "@vscode/extension-telemetry";
import { TelemetryReporter, ConfigFolderName } from "@microsoft/teamsfx-api";
import { anonymizeFilePaths } from "../utils/fileSystemUtils";
import { getPackageVersion } from "../utils/telemetryUtils";
import { TelemetryProperty } from "./extTelemetryEvents";
import {
  Correlator,
  featureFlagManager,
  FeatureFlags,
  getProjectMetadata,
} from "@microsoft/teamsfx-core";
import { configure, getLogger, Logger } from "log4js";
import { workspaceUri } from "../globalVariables";
import VSCodeLogger from "../commonlib/log";

const TelemetryTestLoggerFile = "telemetryTest.log";

/**
 *  VSCode telemetry reporter used by fx-core.
 *  Usage:
 *    let reporter = new VSCodeTelemetryReporter(key, extensionVersion, extensionId)
 *  Illustrate:
 *    key = <'the application insights key'>, 'aiKey' in package.json
 *    extensionVersion = '<extension version>', extension version will be reported as a property with each event
 *    extensionId = '<your extension unique name>', all events will be prefixed with this event name. eg: 'extensionId/eventname'
 */
export class VSCodeTelemetryReporter extends vscode.Disposable implements TelemetryReporter {
  private readonly reporter: Reporter;
  private readonly extVersion: string;
  private readonly logger: Logger | undefined;
  private readonly testFeatureFlag: boolean;

  private sharedProperties: { [key: string]: string } = {};

  constructor(key: string, extensionVersion: string, extensionId: string, reporter?: Reporter) {
    super(async () => await this.reporter.dispose());
    this.reporter = reporter ?? new Reporter(extensionId, extensionVersion, key, true);
    this.extVersion = getPackageVersion(extensionVersion);
    this.testFeatureFlag = featureFlagManager.getBooleanValue(FeatureFlags.TelemetryTest);
    if (this.testFeatureFlag) {
      const logFile = path.join(os.homedir(), `.${ConfigFolderName}`, TelemetryTestLoggerFile);
      configure({
        appenders: { everything: { type: "file", filename: logFile } },
        categories: { default: { appenders: ["everything"], level: "debug" } },
      });
      this.logger = getLogger("TelemTest");
    }
  }

  addSharedProperty(name: string, value?: string): void {
    this.sharedProperties[name] = value ?? "";
  }

  logTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    this.logger?.debug(eventName, properties, measurements);
  }

  logTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {
    this.logger?.debug(eventName, properties, measurements, errorProps);
  }

  logTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    this.logger?.debug(error, properties, measurements);
  }

  sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {
    if (!properties) {
      properties = { ...this.sharedProperties };
    } else {
      properties = { ...this.sharedProperties, ...properties };
    }

    this.checkAndOverwriteSharedProperty(properties);
    if (properties[TelemetryProperty.CorrelationId] == undefined) {
      properties[TelemetryProperty.CorrelationId] = Correlator.getId();
    }

    const featureFlags = featureFlagManager.listEnabled();
    properties[TelemetryProperty.FeatureFlags] = featureFlags ? featureFlags.join(";") : "";

    if (TelemetryProperty.ErrorMessage in properties) {
      properties[TelemetryProperty.ErrorMessage] = anonymizeFilePaths(
        properties[TelemetryProperty.ErrorMessage]
      );
    }

    if (TelemetryProperty.ErrorStack in properties) {
      properties[TelemetryProperty.ErrorStack] = anonymizeFilePaths(
        properties[TelemetryProperty.ErrorStack]
      );
    }

    if (this.testFeatureFlag) {
      this.logTelemetryErrorEvent(eventName, properties, measurements, errorProps);
    } else {
      this.reporter.sendTelemetryErrorEvent(eventName, properties, measurements);
      void VSCodeLogger.debug(
        `sendTelemetryErrorEvent ===> ${eventName}, properties: ${JSON.stringify(
          properties
        )}, measurements: ${JSON.stringify(measurements)}`
      );
    }
  }

  sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    if (!properties) {
      properties = { ...this.sharedProperties };
    } else {
      properties = { ...this.sharedProperties, ...properties };
    }

    this.checkAndOverwriteSharedProperty(properties);
    if (properties[TelemetryProperty.CorrelationId] == undefined) {
      // deactivate event will set correlation id and should not be overwritten
      properties[TelemetryProperty.CorrelationId] = Correlator.getId();
    }

    const featureFlags = featureFlagManager.listEnabled();
    properties[TelemetryProperty.FeatureFlags] = featureFlags ? featureFlags.join(";") : "";

    if (this.testFeatureFlag) {
      this.logTelemetryEvent(eventName, properties, measurements);
    } else {
      this.reporter.sendTelemetryEvent(eventName, properties, measurements);
      void VSCodeLogger.debug(
        `sendTelemetryEvent ===> ${eventName}, properties: ${JSON.stringify(
          properties
        )}, measurements: ${JSON.stringify(measurements)}`
      );
    }
  }

  sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    if (!properties) {
      properties = { ...this.sharedProperties };
    } else {
      properties = { ...this.sharedProperties, ...properties };
    }

    this.checkAndOverwriteSharedProperty(properties);
    properties[TelemetryProperty.CorrelationId] = Correlator.getId();

    const featureFlags = featureFlagManager.listEnabled();
    properties[TelemetryProperty.FeatureFlags] = featureFlags ? featureFlags.join(";") : "";

    if (this.testFeatureFlag) {
      this.logTelemetryException(error, properties, measurements);
    } else {
      this.reporter.sendTelemetryException(error, properties, measurements);
    }
  }

  async dispose() {
    await this.reporter.dispose();
  }

  private checkAndOverwriteSharedProperty(properties: { [p: string]: string }) {
    if (!properties[TelemetryProperty.ProjectId]) {
      const fixedProjectSettings = getProjectMetadata(workspaceUri?.fsPath);

      if (fixedProjectSettings?.projectId) {
        properties[TelemetryProperty.ProjectId] = fixedProjectSettings?.projectId;
        this.sharedProperties[TelemetryProperty.ProjectId] = fixedProjectSettings?.projectId;
      }
    }
  }
}
