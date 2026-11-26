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
      "description": "@assertion the authenticated service connection card in the owned Copilot chat shows an enabled Allow button alongside Cancel for connecting to {{text:appName}} for this and future requests. This is capability consent in Copilot, not the facade's Allow this client page.",
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
      "description": "Click the \"Allow\" button alongside Cancel in the Microsoft 365 Copilot connection card for {{text:appName}} to grant access for this and future requests, not the facade's \"Allow this client\" button.",
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
    }
  ]
}
