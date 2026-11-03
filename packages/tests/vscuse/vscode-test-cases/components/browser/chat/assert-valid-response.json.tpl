{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "id": "assertChatValidResponse",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_assertChatValidResponse_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion agent has valid response.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:browser",
        "action:assert-chat-response",
        "expectation:valid-response",
        "step_retry_timeout: 120"
      ]
    }
  ]
}
