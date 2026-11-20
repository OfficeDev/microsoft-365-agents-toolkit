{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "hostSurface": "copilot",
    "id": "revealLatestCopilotResponse",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_revealLatestCopilotResponse_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 627,
        "y": 591
      },
      "description": "Click the circular down-arrow button above the message composer to scroll to the latest Copilot response and reveal the result hidden behind the composer.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": ["dhash:627:591:16:5:2c2d2d0e6d29934e"],
      "postconditions": [],
      "tags": [
        "component:browser",
        "host_surface:copilot",
        "action:reveal-latest-response",
        "evidence:screenshot_20261109_062712_187"
      ]
    }
  ]
}
