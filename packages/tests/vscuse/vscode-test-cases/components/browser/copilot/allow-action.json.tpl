{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "hostSurface": "copilot",
    "id": "allowCopilotAction",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_allowCopilotAction_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Copilot action-consent Allow button is visible.",
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
      "step_id": "step_allowCopilotAction_focusReply_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {
        "keys": "shift+tab"
      },
      "description": "Move keyboard focus from the Copilot composer into the preceding conversation reply.",
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
        "action:reveal-consent-result",
        "recording:r_1111_045640"
      ]
    },
    {
      "step_id": "step_allowCopilotAction_scrollReply_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {
        "keys": "ctrl+end"
      },
      "description": "Scroll the focused Copilot conversation to the latest reply so its result is visible above the composer.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_allowCopilotAction_focusReply_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "action:reveal-consent-result",
        "recording:r_1111_045640"
      ]
    },
    {
      "step_id": "step_allowCopilotAction_assertDismissed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion Copilot has continued after the Allow action-consent card with a subsequent assistant reply or a service sign-in request. The original Allow card may remain in chat history; its presence alone is not failure. An unchanged consent card without subsequent assistant activity does not satisfy this assertion.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_allowCopilotAction_scrollReply_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "exit_state:action-consent-dismissed",
        "step_retry_timeout: 30"
      ]
    }
  ]
}
