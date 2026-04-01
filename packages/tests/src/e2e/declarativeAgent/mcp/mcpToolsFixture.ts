// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";

/**
 * Minimal MCP tools fixture for E2E tests.
 * Written to a temp file so tests don't depend on local environment.
 */
const mcpToolsContent = {
  tools: [
    {
      name: "microsoft_docs_search",
      description:
        "Search official Microsoft documentation to find relevant content.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            description: "A query about Microsoft products or services",
            type: "string",
          },
        },
        required: ["query"],
      },
    },
    {
      name: "microsoft_docs_fetch",
      description:
        "Fetch and convert a Microsoft Learn documentation page to markdown.",
      inputSchema: {
        type: "object",
        properties: {
          url: {
            description: "URL of the Microsoft documentation page",
            type: "string",
          },
        },
        required: ["url"],
      },
    },
  ],
};

export const mcpToolsFilePath = path.join(
  os.tmpdir(),
  "mcp-e2e-test-tools.json"
);

export async function writeMCPToolsFixture(): Promise<void> {
  await fs.writeJSON(mcpToolsFilePath, mcpToolsContent, { spaces: 2 });
}

export async function removeMCPToolsFixture(): Promise<void> {
  await fs.remove(mcpToolsFilePath);
}
