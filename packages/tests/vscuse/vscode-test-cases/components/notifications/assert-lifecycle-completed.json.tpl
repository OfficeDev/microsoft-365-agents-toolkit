{
  "component": {
    "version": 1,
    "uiSurface": "notifications",
    "id": "assertLifecycleCompleted",
    "parameters": [
      "instanceSuffix",
      "operationTitle",
      "notificationText",
      "retryTimeout"
    ]
  },
  "steps": [
    {
      "step_id": "step_assertLifecycleCompleted_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the current {{text:operationTitle}} operation has finished successfully: a visible Visual Studio Code notification contains the literal text {{text:notificationText}}, there is no active {{text:operationTitle}} progress notification, and there is no failure or cancellation notification for this operation. An older success while another {{text:operationTitle}} operation is still running does not satisfy this assertion.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:notifications",
        "entry_state:lifecycle-result-visible",
        "assertion:lifecycle-completed",
        "step_retry_timeout: {{text:retryTimeout}}"
      ]
    }
  ]
}
