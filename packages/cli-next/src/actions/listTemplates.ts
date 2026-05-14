// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateRegistry, TemplateDescriptor } from "@microsoft/teamsfx-core-next";

export interface TemplateRow {
  id: string;
  name: string;
  category: string;
  languages: string;
}

/**
 * List all registered templates as table-ready rows.
 */
export function listTemplatesAction(registry: TemplateRegistry): TemplateRow[] {
  return registry
    .list()
    .sort((a, b) => a.id.localeCompare(b.id))
    .map((d: TemplateDescriptor) => ({
      id: d.id,
      name: d.name,
      category: d.category,
      languages: d.languages.join(", "),
    }));
}
