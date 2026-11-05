{
  "component": {
    "version": 1,
    "uiSurface": "notifications",
    "id": "expandGrantedSharingNotification",
    "parameters": ["instanceSuffix", "notificationPrefix"],
    "recording": {
      "recordingId": "r_1104_093000",
      "stepId": "step_381d5f43",
      "preconditions": [
        "dhash:614:680:16:5:51c4ae2b28ecca94",
        "dhash:614:680:96:5:a8a855a2b254a6ae",
        "dhash:614:680:0:10:a4a438a0b0383233"
      ]
    }
  },
  "steps": [
    {
      "step_id": "step_expandSharingNotification_target_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 614, "y": 680 },
      "description": "Click the notification message starting with {{text:notificationPrefix}} in the VS Code notification center. Click its message text, not the survey notification or any action buttons.",
      "timeout": 30,
      "continue_on_error": "false",
      "preconditions": [],
      "tags": [
        "component:notifications",
        "action:expand-sharing-notification",
        "ocr:true"
      ]
    },
    {
      "step_id": "step_expandSharingNotification_focus_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the notification message starting with {{text:notificationPrefix}} is focused or highlighted in the VS Code notification center. The survey or another notification does not satisfy this assertion.",
      "timeout": 30,
      "continue_on_error": "false",
      "tags": [
        "component:notifications",
        "action:expand-sharing-notification",
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_expandSharingNotification_expand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "right" },
      "description": "Expand the focused target notification with Right so its complete recipient text is visible.",
      "timeout": 30,
      "continue_on_error": "false",
      "tags": ["component:notifications", "action:expand-sharing-notification"]
    }
  ]
}
