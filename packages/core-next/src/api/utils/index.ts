// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserInteraction } from "../qm/ui";
import { CryptoProvider } from "./crypto";
import { ExpServiceProvider } from "./exp";
import { LogProvider } from "./log";
import { TokenProvider } from "./login";
import { TelemetryReporter } from "./telemetry";
import { TreeProvider } from "./tree";

export * from "./login";
export * from "./log";
export * from "./telemetry";
export * from "./tree";
export * from "./crypto";
export * from "./exp";

export interface Tools {
  logProvider: LogProvider;
  tokenProvider: TokenProvider;
  telemetryReporter?: TelemetryReporter;
  treeProvider?: TreeProvider;
  ui: UserInteraction;
  cryptoProvider?: CryptoProvider;
  expServiceProvider?: ExpServiceProvider;
}
