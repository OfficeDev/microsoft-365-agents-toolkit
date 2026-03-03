// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface Template {
  id: string;
  name: string;
  alias?: string;
  displayName?: string;
  language: "typescript" | "javascript" | "csharp" | "python" | "none" | "common";
  description: string;
  link?: string;
}
