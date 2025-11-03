{
  "servers": {
{{#IsLocalMCP}}
    "{{MCPLocalServerName}}": {
      "type": "stdio",
      "command": "odr.exe",
      "args": [
        "mcp",
        "--proxy",
        "{{MCPLocalServerIdentifier}}"
      ]
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