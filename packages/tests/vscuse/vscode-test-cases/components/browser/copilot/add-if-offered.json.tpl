{
  "component": {
    "version": 1,
    "uiSurface": "browser",
    "hostSurface": "copilot",
    "id": "addAgentIfOffered",
    "parameters": ["instanceSuffix", "scriptBase64"]
  },
  "steps": [
    {
      "step_id": "step_addAgentIfOffered_{{text:instanceSuffix}}",
      "agent": "code",
      "tool": "",
      "parameters": {
        "sample": "=== Generated Script ===\nLanguage: bash\n\n```bash\nset -euo pipefail\nprintf '%s' '{{text:scriptBase64}}' | base64 -d | /usr/bin/python3 - '${{var:app_name}}local'\n```"
      },
      "description": "@code execute the supplied bash script exactly as written in the vscuse container. Do not regenerate the encoded script or click by coordinates. It clicks Add only if offered for the current Copilot agent; an absent installation prompt is normal. Require exit code 0 and the 'Copilot optional Add:' outcome.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:browser", "host_surface:copilot", "action:add-agent-if-offered", "step_retry_timeout: 0"]
    }
  ]
}
