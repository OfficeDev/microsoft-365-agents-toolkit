// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface Template {
  id: string; // internal unique identifier for scaffolding a template folder
  name: string; // unique in UI entry except language and also used in telemetry
  language: "typescript" | "javascript" | "csharp" | "python" | "none" | "common";
  displayName?: string; // used by CLI to show a friendly name
  description: string;
  link?: string;
}
