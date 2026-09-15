{
  "component": {
    "version": 1,
    "uiSurface": "authentication",
    "id": "signOutM365",
    "parameters": [
      "instanceSuffix"
    ]
  },
  "steps": [
    {
      "step_id": "step_signOutM365_open_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "f1"
      },
      "description": "Open the VS Code Command Palette after the local debug session stopped.",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_signOutM365_filter_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "Microsoft 365 Agents: Accounts"
      },
      "description": "Filter for the Microsoft 365 Agents Accounts command.",
      "depends_on": [
        "step_signOutM365_open_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_signOutM365_assertCommands_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Command Palette shows Microsoft 365 Agents Toolkit: Focus on Accounts View as the highlighted first result and Microsoft 365 Agents: Accounts as the second selectable result.",
      "depends_on": [
        "step_signOutM365_filter_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_signOutM365_selectCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "down"
      },
      "description": "Highlight the second result, Microsoft 365 Agents: Accounts.",
      "depends_on": [
        "step_signOutM365_assertCommands_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_signOutM365_execute_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Execute the highlighted Accounts command.",
      "depends_on": [
        "step_signOutM365_selectCommand_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_signOutM365_assertAccounts_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the account picker lists Sign out of Microsoft 365 for the signed-in account.",
      "depends_on": [
        "step_signOutM365_execute_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_signOutM365_filterAccount_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "Sign out of Microsoft 365"
      },
      "description": "Filter the account picker to Sign out of Microsoft 365.",
      "depends_on": [
        "step_signOutM365_assertAccounts_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_signOutM365_assertSelection_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion Sign out of Microsoft 365 is the highlighted account picker option.",
      "depends_on": [
        "step_signOutM365_filterAccount_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_signOutM365_chooseAccount_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Select Sign out of Microsoft 365.",
      "depends_on": [
        "step_signOutM365_assertSelection_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_signOutM365_assertConfirmation_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft 365 sign-out confirmation is visible with Sign out as its focused primary action.",
      "depends_on": [
        "step_signOutM365_chooseAccount_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_signOutM365_confirm_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Confirm signing out of the original Microsoft 365 account.",
      "depends_on": [
        "step_signOutM365_assertConfirmation_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    }
  ]
}
