{
  "component": {
    "version": 1,
    "id": "connectCopilotOAuth",
    "parameters": ["instanceSuffix", "appName"]
  },
  "steps": [
    {
      "step_id": "step_connectCopilotOAuth_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the current Copilot turn requires authentication and shows an enabled Sign in to {{text:appName}} button alongside Cancel. This is the OAuth connection card, not an Allow capability-consent card.",
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
        "entry_state:oauth-connection",
        "step_retry_timeout:120"
      ]
    },
    {
      "step_id": "step_connectCopilotOAuth_click_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 386, "y": 367 },
      "description": "Click the Sign in to {{text:appName}} button in the current Microsoft 365 Copilot turn to start the registered OAuth connection.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_connectCopilotOAuth_assert_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:386:367:16:5:c913ceda925ac920",
        "dhash:386:367:96:5:0000004444000400"
      ],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "action:connect-oauth",
        "ocr:true",
        "source_recording:r_1022_133108"
      ]
    }
  ]
}
