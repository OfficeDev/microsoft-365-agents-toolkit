{
  "component": {
    "version": 1,
    "id": "assertCopilotOAuthReturn",
    "parameters": ["instanceSuffix", "appName"]
  },
  "steps": [
    {
      "step_id": "step_assertCopilotOAuthReturn_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion Microsoft 365 Copilot is active again with the owned agent {{text:appName}} and its message composer visible. No external authentication popup, Microsoft account-picker page, or pending Sign in to {{text:appName}} card is visible. Popup closure alone does not establish a completed tool invocation.",
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
        "assertion:copilot-oauth-return",
        "step_retry_timeout:120"
      ]
    },
    {
      "step_id": "step_assertCopilotOAuthReturn_clearHover_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "esc" },
      "description": "Press Escape in the active Copilot chat to dismiss the sidebar hover tooltip without dismissing the inline connection-consent card.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_assertCopilotOAuthReturn_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "action:clear-oauth-return-hover",
        "source_recording:r_1022_220909"
      ]
    }
  ]
}
