// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DepsTelemetry } from "./depsTelemetry";
import { SystemError, TelemetryReporter, UserError } from "@microsoft/teamsfx-api";
import os from "os";
import { DepsCheckerEvent, TelemetryMessurement } from "./constant";
import { TelemetryProperty, telemetryUtils } from "../../common/telemetry";
import { TelemetryMeasurement } from "../utils/depsChecker/common";

export class CoreDepsTelemetryAdapter implements DepsTelemetry {
  private readonly _telemetryComponentType = "core:debug:envchecker";
  private readonly _telemetryReporter: TelemetryReporter;

  constructor(telemetryReporter: TelemetryReporter) {
    this._telemetryReporter = telemetryReporter;
  }

  public sendEvent(
    eventName: DepsCheckerEvent,
    properties: { [key: string]: string } = {},
    timecost?: number
  ): void {
    this.addCommonProps(properties);
    const measurements: { [p: string]: number } = {};

    if (timecost) {
      measurements[TelemetryMessurement.completionTime] = timecost;
    }

    this._telemetryReporter.sendTelemetryEvent(eventName, properties, measurements);
  }

  public async sendEventWithDuration(
    eventName: DepsCheckerEvent,
    action: () => Promise<void>
  ): Promise<void> {
    const start = Date.now();
    await action();
    // use seconds instead of milliseconds
    const timecost = Number(((Date.now() - start) / 1000).toFixed(2));
    this.sendEvent(eventName, {}, timecost);
  }

  public sendUserErrorEvent(eventName: DepsCheckerEvent, errorMessage: string): void {
    const properties: { [key: string]: string } = {};
    const error = new UserError(this._telemetryComponentType, eventName, errorMessage);
    telemetryUtils.fillInErrorProperties(properties, error);
    this._telemetryReporter.sendTelemetryErrorEvent(eventName, {
      ...this.addCommonProps(),
      ...properties,
    });
  }

  public sendSystemErrorEvent(
    eventName: DepsCheckerEvent,
    errorMessage: string,
    errorStack: string
  ): void {
    const properties: { [key: string]: string } = {};
    const error = new SystemError(
      this._telemetryComponentType,
      eventName,
      `errorMsg=${errorMessage},errorStack=${errorStack}`
    );
    error.stack = errorStack;
    telemetryUtils.fillInErrorProperties(properties, error);
    this._telemetryReporter.sendTelemetryErrorEvent(eventName, {
      ...this.addCommonProps(),
      ...properties,
    });
  }

  private addCommonProps(properties: { [key: string]: string } = {}): { [key: string]: string } {
    properties[TelemetryMeasurement.OSArch] = os.arch();
    properties[TelemetryMeasurement.OSRelease] = os.release();
    properties[TelemetryProperty.Component] = this._telemetryComponentType;
    return properties;
  }
}
