import { createServer } from "../src/server";
import * as fetcherModule from "../src/fetcher";
import { SchemaType } from "../src/fetcher";

describe("MCP Server", () => {
  let fetchSchemaSpy: jest.SpyInstance;

  // Setup and teardown
  beforeEach(() => {
    // Mock the fetchSchema function to avoid actual HTTP requests
    fetchSchemaSpy = jest.spyOn(fetcherModule, "fetchSchema").mockResolvedValue(
      JSON.stringify({
        schema_url: "https://example.com/schema.json",
        content: { test: "content" },
      })
    );
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  describe("createServer", () => {
    it("should create an MCP server with the correct configuration", () => {
      const server = createServer();

      // We can't directly access name and version properties as they're not exposed in the API
      // Instead, we can test that the server instance was created without errors
      expect(server).toBeInstanceOf(Object);
      expect(server).toHaveProperty("server");
      expect(typeof server.connect).toBe("function");
    });

    it("should register the get_schema tool", () => {
      const server = createServer();

      // Since we can't directly access tools property, we can test the functionality
      // by checking that the server's internal structure has our tool
      // We'll use a workaround to test that the tool exists
      const serverAny = server as any;

      // Check if _registeredTools exists and has our get_schema tool
      expect(serverAny._registeredTools).toBeDefined();
      expect(serverAny._registeredTools).toHaveProperty("get_schema");
    });
  });

  describe("get_schema tool", () => {
    it("should call fetchSchema with the correct parameters", async () => {
      // Set up a Transport mock to test the tool functionality
      const mockRequest = {
        params: {
          schema_name: "app_manifest" as SchemaType,
          schema_version: "v1.16",
        },
      };

      // Create server
      const server = createServer();
      const serverAny = server as any;

      // Access the tool callback directly from _registeredTools
      const toolCallback = serverAny._registeredTools["get_schema"].callback;

      // Call the tool callback directly
      await toolCallback(mockRequest.params, { request: mockRequest });

      // Verify fetchSchema was called with the correct parameters
      expect(fetchSchemaSpy).toHaveBeenCalledTimes(1);
      expect(fetchSchemaSpy).toHaveBeenCalledWith("app_manifest", "v1.16");
    });

    it("should return schema content in the response", async () => {
      const expectedSchema = JSON.stringify({
        schema_url: "https://example.com/test.json",
        content: { test: "schema content" },
      });

      // Configure the mock to return the expected schema
      fetchSchemaSpy.mockResolvedValue(expectedSchema);

      // Create server
      const server = createServer();
      const serverAny = server as any;

      // Access the tool callback directly
      const toolCallback = serverAny._registeredTools["get_schema"].callback;

      // Call the tool callback
      const response = await toolCallback(
        {
          schema_name: "api_plugin_manifest" as SchemaType,
          schema_version: "v1.0",
        },
        {}
      );

      // Verify the response contains the expected schema
      expect(response).toHaveProperty("content");
      expect(Array.isArray(response.content)).toBe(true);
      expect(response.content[0]).toHaveProperty("type", "text");
      expect(response.content[0]).toHaveProperty("text", expectedSchema);
    });

    it("should handle different schema types correctly", async () => {
      // Create server
      const server = createServer();
      const serverAny = server as any;

      // Access the tool callback directly
      const toolCallback = serverAny._registeredTools["get_schema"].callback;

      // Test with declarative_agent_manifest
      await toolCallback(
        {
          schema_name: "declarative_agent_manifest" as SchemaType,
          schema_version: "v1.0",
        },
        {}
      );

      expect(fetchSchemaSpy).toHaveBeenCalledWith("declarative_agent_manifest", "v1.0");

      // Clear the mock's history for the next call
      jest.clearAllMocks();

      // Test with api_plugin_manifest
      await toolCallback(
        {
          schema_name: "api_plugin_manifest" as SchemaType,
          schema_version: "v1.0",
        },
        {}
      );

      expect(fetchSchemaSpy).toHaveBeenCalledWith("api_plugin_manifest", "v1.0");
    });

    it("should propagate errors from fetchSchema", async () => {
      const errorMessage = "Failed to fetch schema";
      fetchSchemaSpy.mockResolvedValue(errorMessage);

      // Create server
      const server = createServer();
      const serverAny = server as any;

      // Access the tool callback directly
      const toolCallback = serverAny._registeredTools["get_schema"].callback;

      // Call the tool callback
      const response = await toolCallback(
        {
          schema_name: "app_manifest" as SchemaType,
          schema_version: "invalid-version",
        },
        {}
      );

      // Verify the response contains the error message
      expect(response).toHaveProperty("content");
      expect(Array.isArray(response.content)).toBe(true);
      expect(response.content[0]).toHaveProperty("type", "text");
      expect(response.content[0]).toHaveProperty("text", errorMessage);
    });
  });
});
