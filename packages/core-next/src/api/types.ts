// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  IBot,
  IComposeExtension,
  IConfigurableTab,
  IStaticTab,
  IWebApplicationInfo,
} from "@microsoft/app-manifest";
import { Platform } from "./constants";

export interface OptionItem {
  id: string;
  label: string;
  description?: string;
  detail?: string;
  data?: unknown;
  /** @deprecated CLI will use `cliName` as display name, and use `id` instead if `cliName` is undefined. */
  cliName?: string;
  groupName?: string;
  buttons?: { iconPath: string; tooltip: string; command: string }[];
}

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
export type Void = {};
export const Void = {};

export interface EnvMeta {
  name: string;
  local: boolean;
  sideloading: boolean;
}

export interface Inputs extends Record<string, any> {
  platform: Platform;
  projectPath?: string;
  projectId?: string;
  nonInteractive?: boolean;
  correlationId?: string;
  agent?: "teams" | "office";
  apiAuthData?: AuthInfo[];
}

export type InputsWithProjectPath = Inputs & { projectPath: string };

export type CreateProjectInputs = Inputs & { "app-name": string; folder: string };

export type DeepReadonly<T> = {
  readonly [P in keyof T]: DeepReadonly<T[P]>;
};

export type MaybePromise<T> = T | Promise<T>;

export interface Settings {
  version: string;
  trackingId: string;
}

export type ManifestCapability =
  | { name: "staticTab"; snippet?: IStaticTab; existingApp?: boolean }
  | { name: "configurableTab"; snippet?: IConfigurableTab; existingApp?: boolean }
  | { name: "Bot"; snippet?: IBot; existingApp?: boolean }
  | { name: "MessageExtension"; snippet?: IComposeExtension; existingApp?: boolean }
  | { name: "WebApplicationInfo"; snippet?: IWebApplicationInfo; existingApp?: boolean };

export interface AuthInfo {
  serverUrl: string;
  authName?: string;
  authType?: "apiKey" | "oauth2";
}

export interface ApiOperation {
  id: string;
  label: string;
  groupName: string;
  data: AuthInfo;
  detail?: string;
}

export interface Warning {
  type: string;
  content: string;
  data?: any;
}
