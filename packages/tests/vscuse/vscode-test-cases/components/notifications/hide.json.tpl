{
  "component": {
    "version": 1,
    "uiSurface": "notifications",
    "id": "hideNotifications",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_hideNotifications_hide_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "escape"
      },
      "description": "Press Escape to hide the VS Code notification center without clearing its notifications.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:notifications", "action:hide"]
    },
    {
      "step_id": "step_hideNotifications_assertHidden_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Visual Studio Code notification center is no longer visible and no notification popup obscures the workbench.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_hideNotifications_hide_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:notifications",
        "action:hide",
        "step_retry_timeout: 30"
      ]
    }
  ]
}
