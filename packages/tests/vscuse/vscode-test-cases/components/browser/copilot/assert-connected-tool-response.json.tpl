{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "hostSurface": "copilot",
    "id": "assertConnectedCopilotToolResponse",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_allowCopilotAction_assertDismissed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the latest assistant turn after the most recent user request is complete and contains a non-empty response with returned remote-tool result data. An Allow card retained on an earlier turn may remain visible and does not indicate pending consent for this request. A consent-only response, pending Sign in or Allow request on the latest turn, loading state, refusal, or error is not completion.",
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
        "exit_state:action-consent-dismissed",
        "step_retry_timeout: 120"
      ]
    }
  ]
}
