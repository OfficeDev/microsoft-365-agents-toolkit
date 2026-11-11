{
  "component": {
    "version": 1,
    "id": "verifyNoAzureSubscriptions",
    "parameters": ["instanceSuffix", "script"]
  },
  "steps": [
    {
      "step_id": "step_verifyNoAzureSubscriptions_{{text:instanceSuffix}}",
      "agent": "code",
      "tool": "",
      "parameters": {
        "sample": {{json:script}}
      },
      "description": "execute the supplied generated bash script exactly as authored in the running vscuse container. It authenticates the dedicated Azure user and verifies that ARM returns zero accessible subscriptions. Require exit code 0 and VSCUSE_NO_AZURE_SUBSCRIPTIONS_VERIFIED. Do not print environment variables, credentials, tokens or HTTP response bodies; any failure stops this case.",
      "content_refs": [],
      "timeout": 180,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "account:azure", "fixture:no-subscriptions"]
    }
  ]
}