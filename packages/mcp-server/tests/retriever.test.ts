// filepath: c:\Users\chagon\dev\m365-agents-toolkit\packages\mcp-server\tests\retriever.test.ts
import { retrieveResource, ResourceType } from "../src/retriever";

describe("Retriever", () => {
  describe("retrieveResource", () => {
    it("should return static guidance for documents resource type", async () => {
      const result = await retrieveResource(
        "documents" as ResourceType,
        "How do I use Microsoft Graph?"
      );

      expect(result).toContain("How do I use Microsoft Graph?");
      expect(result).toContain("Learn MCP server");
      expect(result).toContain("https://learn.microsoft.com/api/mcp");
      expect(result).toContain("microsoft_docs_search");
      expect(result).toContain("microsoft_docs_fetch");
    });

    it("should return static guidance for samples resource type", async () => {
      const result = await retrieveResource("samples" as ResourceType, "Sample query");

      expect(result).toContain("Sample query");
      expect(result).toContain("Learn MCP server");
      expect(result).toContain("microsoft_code_sample_search");
    });

    it("should return static guidance for issues resource type", async () => {
      const result = await retrieveResource("issues" as ResourceType, "Test query");

      expect(result).toContain("Test query");
      expect(result).toContain("GitHub MCP server");
      expect(result).toContain("https://api.githubcopilot.com/mcp/");
      expect(result).toContain("OfficeDev/microsoft-365-agents-toolkit");
    });

    it("should return static guidance for code resource type", async () => {
      const result = await retrieveResource("code" as ResourceType, "Teams message extension");

      expect(result).toContain("Teams message extension");
      expect(result).toContain("Learn MCP server");
      expect(result).toContain("microsoft_code_sample_search");
      expect(result).toContain("@microsoft/teams-ai");
    });

    it("should include the user question in the response", async () => {
      const question = "How do I create a Teams app?";
      const result = await retrieveResource("documents" as ResourceType, question);

      expect(result).toContain(`User question: "${question}"`);
    });

    it("should properly handle all resource types", async () => {
      const resourceTypes: ResourceType[] = ["documents", "samples", "issues", "code"];

      for (const resourceType of resourceTypes) {
        const result = await retrieveResource(resourceType, `Query for ${resourceType}`);
        expect(result).toContain(`Query for ${resourceType}`);
        expect(typeof result).toBe("string");
        expect(result.length).toBeGreaterThan(0);
      }
    });
  });
});
