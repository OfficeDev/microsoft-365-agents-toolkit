{
  "component": {
    "version": 1,
    "id": "assertToolCalled",
    "parameters": ["instanceSuffix", "expectedTool"]
  },
  "steps": [
    {
      "step_id": "step_assertToolCalled_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the current completed assistant response contains a returned result for the requested remote tool {{text:expectedTool}}. Returned tool-specific result data, including inline JSON, is sufficient evidence. Do not require a separate tool execution marker, citation, trace panel, debugger information, or developer mode. A user prompt, tool description, refusal, loading state, or promise to call without a result does not satisfy this assertion.",
      "content_refs": [],
      "timeout": 60,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:chat",
        "assertion:tool-called",
        "step_retry_timeout:120"
      ]
    }
  ]
}
