{
  "servers": {
{{#IsLocalMCP}}
    "{{MCPLocalServerName}}": {
      "type": "local",
      "endpoint": "{{MCPLocalServerIdentifier}}"
    }
{{/IsLocalMCP}}
{{^IsLocalMCP}}
    "{{ServerName}}": {
      "type": "remote",
      "url": "{{MCPForDAServerUrl}}"
    }
{{/IsLocalMCP}}
  }
}