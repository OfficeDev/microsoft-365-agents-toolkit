{
  "component": {
    "version": 1,
    "uiSurface": "debug",
    "id": "assertLaunchCancelled",
    "parameters": [
      "instanceSuffix"
    ]
  },
  "steps": [
    {
      "step_id": "step_assertLaunchCancelled_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion Visual Studio Code's project editor is visible after cancelling local launch. The Microsoft 365 environment mismatch dialog is closed. No Microsoft 365 sign-out confirmation or sign-in prompt is open, and no Teams Chrome debug window is in front.",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:60"
      ]
    }
  ]
}
