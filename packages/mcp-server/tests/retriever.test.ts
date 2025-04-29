// filepath: c:\Users\chagon\dev\m365-agents-toolkit\packages\mcp-server\tests\retriever.test.ts
import { retrieveResource, ResourceType } from "../src/retriever";

// Mock fetch
global.fetch = jest.fn();

describe("Retriever", () => {
  // Store the original environment variables
  const originalEnv = process.env;

  beforeEach(() => {
    // Clear all mocks between tests
    jest.clearAllMocks();
    // Reset the environment before each test
    process.env = { ...originalEnv };
    // Mock fetch to return a successful response by default
    (global.fetch as jest.Mock).mockResolvedValue({
      ok: true,
      json: async () => ({ data: "Mocked resource content" }),
    });
  });

  afterEach(() => {
    // Restore the original environment
    process.env = originalEnv;
  });

  describe("retrieveResource", () => {
    it("should retrieve resources from the API service with the correct parameters", async () => {
      // Set environment variables
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "mock-token";
      process.env.RETRIEVER_API_ENDPOINT = "https://example.com/api/retriever";

      // Call the function with test parameters
      const result = await retrieveResource(
        "documents" as ResourceType,
        "How do I use Microsoft Graph?"
      );

      // Verify fetch was called with the right parameters
      expect(global.fetch).toHaveBeenCalledTimes(1);
      expect(global.fetch).toHaveBeenCalledWith("https://example.com/api/retriever", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer mock-token",
        },
        body: JSON.stringify({
          resource_type: "documents",
          question: "How do I use Microsoft Graph?",
        }),
      });

      // Verify the result
      expect(result).toBe("Mocked resource content");
    });

    it("should return an error if no API endpoint is provided", async () => {
      // Set only the token, no endpoint
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "mock-token";
      process.env.RETRIEVER_API_ENDPOINT = "";

      // Call the function
      const result = await retrieveResource("samples" as ResourceType, "Sample query");

      // Verify that fetch was not called
      expect(global.fetch).not.toHaveBeenCalled();

      // Verify the error message
      expect(result).toBe(
        "RETRIEVER_API_ENDPOINT is not set, set it in your MCP server configuration to use this tool."
      );
    });

    it("should return an error message if no token is provided", async () => {
      // Ensure GITHUB_PERSONAL_ACCESS_TOKEN is not set
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "";

      // Call the function
      const result = await retrieveResource("issues" as ResourceType, "Error query");

      // Verify no fetch call was made
      expect(global.fetch).not.toHaveBeenCalled();

      // Verify error message
      expect(result).toBe(
        "GITHUB_PERSONAL_ACCESS_TOKEN is not set, set it in your MCP server configuration to use this tool."
      );
    });

    it("should handle HTTP errors from the API service", async () => {
      // Set environment variables
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "mock-token";
      process.env.RETRIEVER_API_ENDPOINT = "https://example.com/api/retriever";

      // Mock fetch to return an error
      (global.fetch as jest.Mock).mockResolvedValue({
        ok: false,
        status: 404,
        statusText: "Not Found",
      });

      // Call the function
      const result = await retrieveResource("documents" as ResourceType, "Test query");

      // Verify fetch was called
      expect(global.fetch).toHaveBeenCalledTimes(1);

      // Verify error message
      expect(result).toBe("Fail to retrieve resource, 404: Not Found");
    });

    it("should handle network or other errors during fetch", async () => {
      // Set environment variables
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "mock-token";
      process.env.RETRIEVER_API_ENDPOINT = "https://example.com/api/retriever";

      // Mock fetch to throw an error
      const error = new Error("Network error");
      (global.fetch as jest.Mock).mockRejectedValue(error);

      // Call the function
      const result = await retrieveResource("documents" as ResourceType, "Test query");

      // Verify fetch was called
      expect(global.fetch).toHaveBeenCalledTimes(1);

      // Verify error message
      expect(result).toBe("Error retrieving resource: Network error");
    });

    it("should handle non-Error objects thrown during fetch", async () => {
      // Set environment variables
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "mock-token";
      process.env.RETRIEVER_API_ENDPOINT = "https://example.com/api/retriever";

      // Mock fetch to throw a non-Error object
      (global.fetch as jest.Mock).mockRejectedValue("Some string error");

      // Call the function
      const result = await retrieveResource("documents" as ResourceType, "Test query");

      // Verify fetch was called
      expect(global.fetch).toHaveBeenCalledTimes(1);

      // Verify error message
      expect(result).toBe("Error retrieving resource: Unknown error occurred");
    });

    it("should properly handle all resource types", async () => {
      // Set environment variables
      process.env.GITHUB_PERSONAL_ACCESS_TOKEN = "mock-token";
      process.env.RETRIEVER_API_ENDPOINT = "https://example.com/api/retriever";

      // Test each resource type
      const resourceTypes: ResourceType[] = ["documents", "samples", "issues"];

      for (const resourceType of resourceTypes) {
        // Clear previous calls before testing each resource type
        jest.clearAllMocks();

        // Call the function for the current resource type
        await retrieveResource(resourceType, `Query for ${resourceType}`);

        // Verify fetch was called with the correct resource type
        expect(global.fetch).toHaveBeenCalledTimes(1);
        const callBody = JSON.parse((global.fetch as jest.Mock).mock.calls[0][1].body);
        expect(callBody.resource_type).toBe(resourceType);
      }
    });
  });
});
