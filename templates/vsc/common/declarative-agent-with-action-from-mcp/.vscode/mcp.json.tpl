{
  "servers": {
{{#IsLocalMCP}}
{{#MCPLocalServers}}
    "{{name}}": {
      "type": "stdio",
      "command": "{{command}}",
      "args": [{{args}}]
    }{{#notLast}},{{/notLast}}
{{/MCPLocalServers}}
{{/IsLocalMCP}}
{{^IsLocalMCP}}
    "{{ServerName}}": {
      "type": "remote",
      "url": "{{MCPForDAServerUrl}}"
    }
{{/IsLocalMCP}}
  }
}