{
  "component": {
    "version": 1,
    "id": "assertChatNonEmptyField",
    "parameters": ["instanceSuffix", "fieldName"]
  },
  "steps": [
    {
      "step_id": "step_assertChatNonEmptyField_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the current completed assistant tool result contains the field {{text:fieldName}} with a non-empty string value. Missing, null, empty and whitespace-only values fail. Any non-empty identifier is acceptable; no specific literal value or JSON whitespace formatting is required. The field name in the user prompt or a tool description alone is not a result.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:chat",
        "action:assert-chat-response",
        "expectation:non-empty-field",
        "step_retry_timeout:120"
      ]
    }
  ]
}
