{
  "component": {
    "version": 1,
    "uiSurface": "notifications",
    "id": "assertCollapsedPrefixAndContains",
    "parameters": [
      "instanceSuffix",
      "collapsedNotificationPrefix",
      "notificationText",
      "retryTimeout"
    ]
  },
  "steps": [
    {
      "step_id": "step_assertNotifications_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the visible Visual Studio Code Notification Center contains both of these notifications: (1) a collapsed yellow warning notification whose visible text begins with the literal prefix {{text:collapsedNotificationPrefix}} and is truncated with an ellipsis; (2) an information notification containing the complete literal text {{text:notificationText}} Notifications with different text, including in-progress notifications, do not satisfy this assertion.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:notifications",
        "action:assert-collapsed-prefix-and-contains",
        "step_retry_timeout: {{text:retryTimeout}}"
      ]
    },
    {
      "step_id": "step_assertNotifications_hide_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 998,
        "y": 343
      },
      "description": "Click the Hide Notifications downward chevron at the far right of the Notification Center header, not a notification's Close button or the Clear All Notifications control.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_assertNotifications_assert_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:998:343:16:5:685596cc69339474",
        "dhash:998:343:96:5:0048b2484836c8c8"
      ],
      "postconditions": [],
      "tags": ["component:notifications", "action:hide"]
    },
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
      "depends_on": ["step_assertNotifications_hide_{{text:instanceSuffix}}"],
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
