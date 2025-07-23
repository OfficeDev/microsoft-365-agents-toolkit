{
  "servers": {
    // This is the MCP server configuration file for VS Code to use.
    // To add your MCP server for the Declarative Agent, add a new entry under the “servers” section. Then click the “Start” button from CodeLens. 
    // To add tools fetched from the MCP server, click the “ATK: Update Action with MCP” button from CodeLens or Command Palette, and then select the wanted tools from the prompt list. 
    "{{ServerName}}": {
			"url": "{{MCPForDAServerUrl}}",
			"type": "http"
		}
  }
}