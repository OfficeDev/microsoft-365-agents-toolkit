{
  "component": {
    "version": 1,
    "id": "invalidPluginProperty",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_invalidPluginProperty_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft 365 Agents Toolkit Output panel identifies ai-plugin.json with a validation error saying property name must be valid and identifies the exact invalid property name_for_humanaaaa. A generic failure or an unrelated validation error does not satisfy this assertion.",
      "tags": [
        "component:checks",
        "operation:validate-manifest",
        "assertion:invalid-plugin-property",
        "step_retry_timeout:120"
      ]
    }
  ]
}
