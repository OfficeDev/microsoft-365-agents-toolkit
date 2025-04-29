// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { fetchSchema, SchemaTypeEnum } from "./fetcher";
import { retrieveResource } from "./retriever";

export function createServer(): McpServer {
  const server = new McpServer({
    name: "m365agentstoolkit-mcp",
    version: "0.1.0",
  });
  server.tool(
    "get_schema",
    'Get the schema for "App manifest", "Declarative agent manifest", "API plugin manifest", use it everytime before understanding, modifying or creating any of these manifest files.',
    {
      schema_name: SchemaTypeEnum.describe("name of schema"),
      schema_version: z.string().describe("version of schema"),
    },
    async ({ schema_name, schema_version }) => {
      const schema = await fetchSchema(schema_name, schema_version);

      return {
        content: [
          {
            type: "text",
            text: schema,
          },
        ],
      };
    }
  );

  server.tool(
    "get_knowledge",
    "Access comprehensive knowledge about Microsoft 365 and Microsoft 365 Copilot development. Use this tool everytime for questions related to Microsoft 365 and Microsoft 365 Copilot.",
    {
      question: z.string().describe("Question to use for knowledge retrieval"),
    },
    async ({ question }) => {
      const result = await retrieveResource("documents", question);

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };
    }
  );

  server.tool(
    "get_samples",
    "Access templates and code samples for Microsoft 365 and Microsoft 365 Copilot development. Use this tool when looking for implementation examples, starter templates, or reference architectures.",
    {
      question: z.string().describe("Query to find relevant samples and templates"),
    },
    async ({ question }) => {
      const result = await retrieveResource("samples", question);

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };
    }
  );

  server.tool(
    "troubleshoot",
    "Access troubleshooting solutions for common Microsoft 365 and Microsoft 365 Copilot development issues. Use this tool when encountering errors, unexpected behaviors, or implementation challenges.",
    {
      question: z.string().describe("Description of the issue or error you're experiencing"),
    },
    async ({ question }) => {
      const result = await retrieveResource("issues", question);

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };
    }
  );

  return server;
}
