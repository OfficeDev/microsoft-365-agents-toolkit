// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Inputs } from "@microsoft/teamsfx-api";
import { telemetryUtils, getHashedEnv } from "@microsoft/teamsfx-core";
import { CliTelemetryReporter } from "../commonlib/telemetry";
import { TelemetryComponentType, TelemetryProperty, TelemetrySuccess } from "./cliTelemetryEvents";

export function makeEnvRelatedProperty(
  projectDir: string,
  inputs: Inputs
): { [key: string]: string } {
  const properties: { [key: string]: string } = {};
  if (inputs.env) {
    properties[TelemetryProperty.Env] = getHashedEnv(inputs.env);
  }
  return properties;
}

class CliTelemetry {
  reporter: CliTelemetryReporter | undefined;
  rootFolder: string | undefined;

  set enable(value: boolean) {
    if (this.reporter) {
      this.reporter.enable = value;
    }
  }

  public withRootFolder(rootFolder: string | undefined): CliTelemetry {
    this.rootFolder = rootFolder;
    this.reporter?.withRootFolder(rootFolder);
    return this;
  }

  public sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    if (!properties) {
      properties = {};
    }
    if (TelemetryProperty.Component in properties === false) {
      properties[TelemetryProperty.Component] = TelemetryComponentType;
    }

    properties[TelemetryProperty.Success] = TelemetrySuccess.Yes;

    this.reporter
      ?.withRootFolder(this.rootFolder)
      .sendTelemetryEvent(eventName, properties, measurements);
  }

  public sendTelemetryErrorEvent(
    eventName: string,
    error: FxError,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number },
    errorProps?: string[]
  ): void {
    if (!properties) {
      properties = {};
    }
    if (TelemetryProperty.Component in properties === false) {
      properties[TelemetryProperty.Component] = TelemetryComponentType;
    }

    telemetryUtils.fillInErrorProperties(properties, error);

    this.reporter
      ?.withRootFolder(this.rootFolder)
      .sendTelemetryErrorEvent(eventName, properties, measurements, errorProps);
  }

  public sendTelemetryException(
    error: Error,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {
    if (!properties) {
      properties = {};
    }
    if (TelemetryProperty.Component in properties === false) {
      properties[TelemetryProperty.Component] = TelemetryComponentType;
    }

    this.reporter
      ?.withRootFolder(this.rootFolder)
      .sendTelemetryException(error, properties, measurements);
  }

  public async flush(): Promise<void> {
    await this.reporter?.flush();
  }
}

export default new CliTelemetry();
