{
  "component": {
    "version": 1,
    "uiSurface": "notifications",
    "id": "assertNotificationsHidden",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_assertNotifications_assertHidden_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Visual Studio Code Notification Center is no longer visible and no notification overlays the workbench.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:notifications",
        "action:assert-hidden",
        "step_retry_timeout: 30"
      ]
    }
  ]
}
