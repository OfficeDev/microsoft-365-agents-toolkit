// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author yuqzho@microsoft.com, Ning Tang
 */

import { ProjectType } from "@microsoft/m365-spec-parser";

export interface TemplateState {
  isYaml: boolean;
  templateName: string;
  url: string;
  isPlugin: boolean;
  type: ProjectType;
}

export const enum telemetryProperties {
  templateName = "template-name",
  generateType = "generate-type",
  isRemoteUrlTelemetryProperty = "remote-url",
  authType = "auth-type",
  isDeclarativeCopilot = "is-declarative-copilot",
}
