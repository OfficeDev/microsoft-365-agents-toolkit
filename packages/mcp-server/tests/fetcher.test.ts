import { fetchSchema, SchemaType, clearSchemaCache } from "../src/fetcher";

// Mock the global fetch function
global.fetch = jest.fn();

describe("fetcher", () => {
  beforeEach(() => {
    // Clear the mocks and cache before each test
    jest.clearAllMocks();
    clearSchemaCache();

    // Setup the mock fetch implementation
    (global.fetch as jest.Mock).mockResolvedValue({
      ok: true,
      json: async () => ({ test: "schema content" }),
    });
  });

  describe("fetchSchema", () => {
    it("should fetch schema from the correct URL", async () => {
      const schemaName: SchemaType = "app_manifest";
      const schemaVersion = "v1.16";

      await fetchSchema(schemaName, schemaVersion);

      const expectedUrl = `https://developer.microsoft.com/json-schemas/teams/v1.16/MicrosoftTeams.schema.json`;
      expect(global.fetch).toHaveBeenCalledWith(expectedUrl);
    });

    it("should return schema with URL and content", async () => {
      const schemaName: SchemaType = "app_manifest";
      const schemaVersion = "v1.16";

      const result = await fetchSchema(schemaName, schemaVersion);
      const parsedResult = JSON.parse(result);

      expect(parsedResult).toHaveProperty("schema_url");
      expect(parsedResult).toHaveProperty("content");
      expect(parsedResult.content).toEqual({ test: "schema content" });
    });

    it("should support different schema types", async () => {
      // Test declarative_agent_manifest
      await fetchSchema("declarative_agent_manifest", "v1.0");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.0/schema.json"
      );

      jest.clearAllMocks();

      // Test api_plugin_manifest
      await fetchSchema("api_plugin_manifest", "v1.0");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://developer.microsoft.com/json-schemas/copilot/plugin/v1.0/schema.json"
      );

      jest.clearAllMocks();

      // Test m365_agents_yaml
      await fetchSchema("m365_agents_yaml", "v1.0");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://aka.ms/teams-toolkit/v1.0/yaml.schema.json"
      );
    });

    it("should handle 'latest' version by using the repository's latest version", async () => {
      // Test app_manifest with 'latest' version
      await fetchSchema("app_manifest", "latest");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://developer.microsoft.com/json-schemas/teams/v1.22/MicrosoftTeams.schema.json"
      );

      jest.clearAllMocks();

      // Test declarative_agent_manifest with 'latest' version
      await fetchSchema("declarative_agent_manifest", "latest");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.4/schema.json"
      );

      jest.clearAllMocks();

      // Test api_plugin_manifest with 'latest' version
      await fetchSchema("api_plugin_manifest", "latest");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.3/schema.json"
      );

      jest.clearAllMocks();

      // Test m365_agents_yaml with 'latest' version
      await fetchSchema("m365_agents_yaml", "latest");
      expect(global.fetch).toHaveBeenCalledWith(
        "https://aka.ms/teams-toolkit/v1.8/yaml.schema.json"
      );
    });

    it("should cache results for repeated requests", async () => {
      // First call should make a fetch request
      await fetchSchema("app_manifest", "v1.16");
      expect(global.fetch).toHaveBeenCalledTimes(1);

      // Second call with same parameters should use the cached result
      await fetchSchema("app_manifest", "v1.16");
      expect(global.fetch).toHaveBeenCalledTimes(1); // Still just one call

      // Different parameters should make a new fetch
      await fetchSchema("app_manifest", "v1.15");
      expect(global.fetch).toHaveBeenCalledTimes(2);
    });

    it("should handle HTTP errors", async () => {
      // Mock a failed response
      (global.fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        status: 404,
      });

      const result = await fetchSchema("app_manifest", "invalid-version");

      expect(result).toContain("Failed fetching schema at version: invalid-version");
      expect(result).toContain("HTTP error with status: 404");
    });

    it("should handle network errors", async () => {
      // Mock a network error
      (global.fetch as jest.Mock).mockRejectedValueOnce(new Error("Network error"));

      const result = await fetchSchema("app_manifest", "v1.16");

      expect(result).toContain("Failed fetching schema at version: v1.16");
      expect(result).toContain("Network error");
    });

    it("should handle unknown schema types", async () => {
      const result = await fetchSchema("unknown_schema_type" as SchemaType, "v1.0");

      expect(result).toContain("Unknown schema name");
    });

    it("should fetch m365_agents_yaml schema correctly", async () => {
      const result = await fetchSchema("m365_agents_yaml", "v1.5");
      const parsedResult = JSON.parse(result);

      expect(global.fetch).toHaveBeenCalledWith(
        "https://aka.ms/teams-toolkit/v1.5/yaml.schema.json"
      );
      expect(parsedResult).toHaveProperty("schema_url");
      expect(parsedResult).toHaveProperty("content");
      expect(parsedResult.schema_url).toBe("https://aka.ms/teams-toolkit/v1.5/yaml.schema.json");
      expect(parsedResult.content).toEqual({ test: "schema content" });
    });
  });
});
