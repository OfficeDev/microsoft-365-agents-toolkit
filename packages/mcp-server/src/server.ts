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

  server.registerTool(
    "get_schema",
    {
      description:
        'Get the schema for "App manifest", "Declarative agent manifest", "API plugin manifest", "M365 agents yaml", use it everytime before understanding, modifying or creating any of these manifest files.',
      inputSchema: {
        schema_name: SchemaTypeEnum.describe("name of schema"),
        schema_version: z
          .string()
          .describe(
            'version of schema in semantic versioning format vX.Y, where X is the major version and Y is the minor version (e.g. v1.0, v1.19, v2.1). Use "latest" if unsure.'
          ),
      },
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

  server.registerTool(
    "get_knowledge",
    {
      description:
        "Get documentation and guidance for Microsoft 365 agents and apps development, including Copilot extensibility, Teams platform, Agents Toolkit, and Agents SDK. Use this tool for any questions related to Microsoft 365 and Microsoft 365 Copilot development.",
      inputSchema: {
        question: z.string().describe("Question to use for knowledge retrieval"),
      },
    },
    ({ question }) => {
      const result = retrieveResource("documents", question);

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

  server.registerTool(
    "get_code_snippets",
    {
      description:
        "Get code snippets, templates, and sample repositories for Microsoft 365 agents and apps development. Covers SDKs including **@microsoft/teams-ai**, **@microsoft/teams-js**, **botbuilder**, **@microsoft/agents-hosting**, **@microsoft/agents-activity**, and **@microsoft/teamsfx**. Use this tool when looking for implementation examples, starter templates, or SDK usage patterns.",
      inputSchema: {
        question: z
          .string()
          .describe(
            "Query to find relevant code snippets related to Microsoft 365 app or agent SDKs"
          ),
      },
    },
    ({ question }) => {
      const result = retrieveResource("code", question);

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

  server.registerTool(
    "troubleshoot",
    {
      description:
        "Find troubleshooting solutions and related issues for Microsoft 365 agents and apps development. Searches across Agents Toolkit, Teams SDK, Teams Samples, and documentation repositories. Use this tool when encountering errors, unexpected behaviors, or implementation challenges.",
      inputSchema: {
        question: z.string().describe("Description of the issue or error you're experiencing"),
      },
    },
    ({ question }) => {
      const result = retrieveResource("issues", question);

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
