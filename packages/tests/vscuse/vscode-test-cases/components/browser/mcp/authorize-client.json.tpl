{
  "component": {
    "version": 1,
    "id": "authorizeMcpClient",
    "parameters": ["instanceSuffix", "mcpEndpoint"]
  },
  "steps": [
    {
      "step_id": "step_authorizeMcpClient_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the active authorization popup is on the facade hosting {{text:mcpEndpoint}} and shows Authorize an MCP client, a Client ID, Callback destination https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect, and an enabled Allow this client button. No invalid_client or authorization error is visible.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "entry_state:mcp-client-consent",
        "step_retry_timeout:120"
      ]
    },
    {
      "step_id": "step_authorizeMcpClient_allow_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 62, "y": 389 },
      "description": "Click Allow this client on the facade's Authorize an MCP client page to approve this dynamically registered client's displayed permissions and callback, then continue to Microsoft Entra authorization.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_authorizeMcpClient_assert_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:62:389:16:5:477655454644bb44",
        "dhash:62:389:96:5:e9c9006969002900"
      ],
      "postconditions": [],
      "tags": [
        "component:browser",
        "action:authorize-mcp-client",
        "ocr:true",
        "source_recording:r_1022_133108"
      ]
    }
  ]
}
