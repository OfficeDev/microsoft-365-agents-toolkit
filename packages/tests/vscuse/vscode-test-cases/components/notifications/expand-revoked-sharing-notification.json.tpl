{
  "component": {
    "version": 1,
    "uiSurface": "notifications",
    "id": "expandRevokedSharingNotification",
    "parameters": ["instanceSuffix", "notificationPrefix"],
    "recording": {
      "recordingId": "r_1104_093603",
      "stepId": "step_f519109d",
      "preconditions": [
        "dhash:620:622:16:5:16b5eee47595c41a",
        "dhash:620:622:96:5:484897686a956266",
        "dhash:620:622:0:10:a4a438a0b0323234"
      ]
    }
  },
  "steps": [
    {
      "step_id": "step_expandSharingNotification_target_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 620, "y": 622 },
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
