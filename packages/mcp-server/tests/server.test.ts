// filepath: c:\Users\chagon\dev\m365-agents-toolkit\packages\mcp-server\tests\server.test.ts
import { createServer } from "../src/server";
import * as fetcherModule from "../src/fetcher";
import * as retrieverModule from "../src/retriever";
import { SchemaType } from "../src/fetcher";

describe("MCP Server", () => {
  let fetchSchemaSpy: jest.SpyInstance;
  let retrieveResourceSpy: jest.SpyInstance;

  // Setup and teardown
  beforeEach(() => {
    // Mock the fetchSchema function to avoid actual HTTP requests
    fetchSchemaSpy = jest.spyOn(fetcherModule, "fetchSchema").mockResolvedValue(
      JSON.stringify({
        schema_url: "https://example.com/schema.json",
        content: { test: "content" },
      })
    );

    // Mock the retrieveResource function to avoid actual API calls
    retrieveResourceSpy = jest
      .spyOn(retrieverModule, "retrieveResource")
      .mockResolvedValue("Mocked resource content");
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

    it("should register the get_knowledge, get_samples, and troubleshoot tools", () => {
      const server = createServer();
      const serverAny = server as any;

      // Check if _registeredTools has our new tools
      expect(serverAny._registeredTools).toHaveProperty("get_knowledge");
      expect(serverAny._registeredTools).toHaveProperty("get_samples");
      expect(serverAny._registeredTools).toHaveProperty("troubleshoot");
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

  describe("knowledge and resource tools", () => {
    it("get_knowledge tool should call retrieveResource with correct parameters", async () => {
      const server = createServer();
      const serverAny = server as any;
      const toolCallback = serverAny._registeredTools["get_knowledge"].callback;

      // Call the tool callback with a question
      await toolCallback({ question: "How do I create a Teams app?" }, {});

      // Verify retrieveResource was called with the correct parameters
      expect(retrieveResourceSpy).toHaveBeenCalledTimes(1);
      expect(retrieveResourceSpy).toHaveBeenCalledWith("documents", "How do I create a Teams app?");
    });

    it("get_samples tool should call retrieveResource with correct parameters", async () => {
      const server = createServer();
      const serverAny = server as any;
      const toolCallback = serverAny._registeredTools["get_samples"].callback;

      // Call the tool callback with a query
      await toolCallback({ question: "Teams message extension sample" }, {});

      // Verify retrieveResource was called with the correct parameters
      expect(retrieveResourceSpy).toHaveBeenCalledTimes(1);
      expect(retrieveResourceSpy).toHaveBeenCalledWith("samples", "Teams message extension sample");
    });

    it("troubleshoot tool should call retrieveResource with correct parameters", async () => {
      const server = createServer();
      const serverAny = server as any;
      const toolCallback = serverAny._registeredTools["troubleshoot"].callback;

      // Call the tool callback with an issue description
      await toolCallback({ question: "App manifest validation error" }, {});

      // Verify retrieveResource was called with the correct parameters
      expect(retrieveResourceSpy).toHaveBeenCalledTimes(1);
      expect(retrieveResourceSpy).toHaveBeenCalledWith("issues", "App manifest validation error");
    });

    it("should return resource content in the response", async () => {
      const mockedContent = "Here is the information you requested about Microsoft 365 development";
      retrieveResourceSpy.mockResolvedValue(mockedContent);

      const server = createServer();
      const serverAny = server as any;
      const toolCallback = serverAny._registeredTools["get_knowledge"].callback;

      // Call the tool callback
      const response = await toolCallback({ question: "How do I use Microsoft Graph?" }, {});

      // Verify the response contains the expected content
      expect(response).toHaveProperty("content");
      expect(Array.isArray(response.content)).toBe(true);
      expect(response.content[0]).toHaveProperty("type", "text");
      expect(response.content[0]).toHaveProperty("text", mockedContent);
    });

    it("should propagate errors from retrieveResource", async () => {
      const errorMessage = "Failed to retrieve resource";
      retrieveResourceSpy.mockResolvedValue(errorMessage);

      const server = createServer();
      const serverAny = server as any;
      const toolCallback = serverAny._registeredTools["get_samples"].callback;

      // Call the tool callback
      const response = await toolCallback({ question: "Invalid query" }, {});

      // Verify the response contains the error message
      expect(response).toHaveProperty("content");
      expect(Array.isArray(response.content)).toBe(true);
      expect(response.content[0]).toHaveProperty("type", "text");
      expect(response.content[0]).toHaveProperty("text", errorMessage);
    });
  });
});
