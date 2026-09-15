{
  "component": {
    "version": 1,
    "uiSurface": "dialog",
    "id": "tenantMismatch",
    "parameters": [
      "instanceSuffix",
      "actionLabel",
      "actionKey"
    ]
  },
  "steps": [
    {
      "step_id": "step_assertTenantMismatch_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft 365 environment mismatch dialog says You're signed in with a Microsoft 365 account that doesn't match this environment. Please sign out and sign in with the correct one. It offers Cancel and Continue, with Continue as the focused primary action.",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:180"
      ]
    },
    {
      "step_id": "step_tenantMismatch_choose_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": {{json:actionKey}}
      },
      "description": "Choose {{text:actionLabel}} in the verified Microsoft 365 environment mismatch dialog.",
      "depends_on": [
        "step_assertTenantMismatch_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    }
  ]
}
