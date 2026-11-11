{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "hostSurface": "copilot",
    "id": "allowConnectedCopilotAction",
    "parameters": ["instanceSuffix", "appName"]
  },
  "steps": [
    {
      "step_id": "step_allowCopilotAction_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the latest assistant turn after the OAuth sign-in request shows an enabled Allow button alongside Cancel for connecting to {{text:appName}}. This is capability consent in Copilot, not the facade's Allow this client page.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "entry_state:action-consent",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_allowCopilotAction_click_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 333,
        "y": 327
      },
      "description": "Click the \"Allow\" button in the Microsoft 365 Copilot chat interface to grant the agent access.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_allowCopilotAction_assert_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "entry_state:action-consent",
        "action:allow",
        "ocr:true"
      ]
    },
    {
      "step_id": "step_allowCopilotAction_assertReady_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the owned Copilot chat is ready for another message: its message composer is visible and enabled, no response is loading or being generated, and no active Sign in request is pending. A retained Allow card from the preceding turn may remain visible; it is not tool-result evidence.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_allowCopilotAction_click_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "exit_state:connected-chat-ready",
        "step_retry_timeout: 120"
      ]
    }
  ]
}
