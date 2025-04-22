import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { fetchSchema, SchemaTypeEnum } from "./fetcher";

export function createServer(): McpServer {
  const server = new McpServer({
    name: "m365agentstoolkit-mcp",
    version: "0.1.0",
  });
  server.tool(
    "get_schema",
    "Get the schema for \"App manifest\", \"Declarative agent manifest\", \"API plugin manifest\", use it everytime before understanding, modifying or creating any of these manifest files.",
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
              text: schema
            }
          ]
        };
    },
  );

  return server;
}
