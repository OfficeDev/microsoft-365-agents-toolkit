// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export interface TemplateConfig {
  id: string;
  components: AtkConfig[];
  features?: Record<string, unknown>;
}

export interface AtkConfig {
  name: string;
}
