{
  "servers": {
{{#IsLocalMCP}}
    "{{MCPLocalServerName}}": {
      "type": "stdio",
      "command": "{{MCPCommand}}",
      "args": [{{MCPArgs}}]
    }
{{/IsLocalMCP}}
{{^IsLocalMCP}}
    "{{ServerName}}": {
      "type": "http",
      "url": "{{MCPForDAServerUrl}}"
    }
{{/IsLocalMCP}}
  }
}