{
  "component": {
    "version": 1,
    "id": "selectSignedInMicrosoftAccount",
    "parameters": ["instanceSuffix", "accountName"]
  },
  "steps": [
    {
      "step_id": "step_selectSignedInMicrosoftAccount_assert_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the active Microsoft sign-in popup shows Pick an account. The Signed in row for {{text:accountName}} has keyboard focus, indicated by its dotted outline. Use another account is not focused.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "entry_state:signed-in-account-picker",
        "step_retry_timeout:120"
      ]
    },
    {
      "step_id": "step_selectSignedInMicrosoftAccount_select_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "enter" },
      "description": "Press Enter to select the focused signed-in Microsoft account {{text:accountName}} and complete this OAuth authorization, not another account.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_selectSignedInMicrosoftAccount_assert_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "action:select-microsoft-account",
        "source_recording:r_1022_210211"
      ]
    }
  ]
}
