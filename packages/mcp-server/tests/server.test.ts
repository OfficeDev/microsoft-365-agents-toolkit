import { expect } from "chai";
import * as sinon from "sinon";
import { createServer } from "../src/server";
import * as fetcherModule from "../src/fetcher";
import { SchemaType } from "../src/fetcher";

describe("MCP Server", () => {
  let fetchSchemaSpy: sinon.SinonStub;

  // Setup and teardown
  beforeEach(() => {
    // Stub the fetchSchema function to avoid actual HTTP requests
    fetchSchemaSpy = sinon.stub(fetcherModule, "fetchSchema").resolves(
      JSON.stringify({
        schema_url: "https://example.com/schema.json",
        content: { test: "content" },
      })
    );
  });

  afterEach(() => {
    sinon.restore();
  });
  describe("createServer", () => {
    it("should create an MCP server with the correct configuration", () => {
      const server = createServer();

      // We can't directly access name and version properties as they're not exposed in the API
      // Instead, we can test that the server instance was created without errors
      expect(server).to.be.an.instanceOf(Object);
      expect(server).to.have.property("server");
      expect(server).to.have.property("connect").that.is.a("function");
    });

    it("should register the get_schema tool", () => {
      const server = createServer();

      // Since we can't directly access tools property, we can test the functionality
      // by checking that the server's internal structure has our tool
      // We'll use a workaround to test that the tool exists
      const serverAny = server as any;

      // Check if _registeredTools exists and has our get_schema tool
      expect(serverAny._registeredTools).to.exist;
      expect(serverAny._registeredTools).to.have.property("get_schema");
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
      expect(fetchSchemaSpy.calledOnce).to.be.true;
      expect(fetchSchemaSpy.calledWith("app_manifest", "v1.16")).to.be.true;
    });

    it("should return schema content in the response", async () => {
      const expectedSchema = JSON.stringify({
        schema_url: "https://example.com/test.json",
        content: { test: "schema content" },
      });

      // Configure the stub to return the expected schema
      fetchSchemaSpy.resolves(expectedSchema);

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
      expect(response).to.have.property("content").that.is.an("array");
      expect(response.content[0]).to.have.property("type", "text");
      expect(response.content[0]).to.have.property("text", expectedSchema);
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

      expect(fetchSchemaSpy.calledWith("declarative_agent_manifest", "v1.0")).to.be.true;

      // Reset the spy to test another schema type
      fetchSchemaSpy.resetHistory();

      // Test with api_plugin_manifest
      await toolCallback(
        {
          schema_name: "api_plugin_manifest" as SchemaType,
          schema_version: "v1.0",
        },
        {}
      );

      expect(fetchSchemaSpy.calledWith("api_plugin_manifest", "v1.0")).to.be.true;
    });

    it("should propagate errors from fetchSchema", async () => {
      const errorMessage = "Failed to fetch schema";
      fetchSchemaSpy.resolves(errorMessage);

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
      expect(response).to.have.property("content").that.is.an("array");
      expect(response.content[0]).to.have.property("type", "text");
      expect(response.content[0]).to.have.property("text", errorMessage);
    });
  });
});
