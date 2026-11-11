{
  "component": {
    "version": 1,
    "uiSurface": "authentication",
    "id": "reauthenticateM365",
    "parameters": [
      "instanceSuffix",
      "accountName",
      "accountPassword"
    ]
  },
  "steps": [
    {
      "step_id": "step_reauthenticateM365_assertSignOut_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft 365 sign-out confirmation dialog is visible with Sign out as its focused primary action.",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:60"
      ]
    },
    {
      "step_id": "step_reauthenticateM365_signOut_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Confirm signing out of the alternate Microsoft 365 account.",
      "depends_on": [
        "step_reauthenticateM365_assertSignOut_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_reauthenticateM365_assertSignIn_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion Microsoft 365 Agents Toolkit requests a Microsoft 365 account with Sign in as the focused primary action.",
      "depends_on": [
        "step_reauthenticateM365_signOut_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:60"
      ]
    },
    {
      "step_id": "step_reauthenticateM365_signIn_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Begin sign-in to restore the original Microsoft 365 account.",
      "depends_on": [
        "step_reauthenticateM365_assertSignIn_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_reauthenticateM365_assertPicker_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft Pick an account page lists the original account {{text:accountName}}.",
      "depends_on": [
        "step_reauthenticateM365_signIn_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "step_retry_timeout:60"
      ]
    },
    {
      "step_id": "step_reauthenticateM365_selectOriginal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 471,
        "y": 453
      },
      "description": "Select the original account from the recorded Microsoft account picker.",
      "depends_on": [
        "step_reauthenticateM365_assertPicker_{{text:instanceSuffix}}"
      ],
      "preconditions": [
        "dhash:471:453:16:5:68765634545aa552",
        "dhash:471:453:96:5:4000628a2d220000",
        "dhash:471:453:0:10:1312f89a9ed5e6e4"
      ],
      "postconditions": [],
      "tags": [
        "precondition_wait_timeout:60"
      ]
    },
    {
      "step_id": "step_reauthenticateM365_assertPassword_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft sign-in page names {{text:accountName}} and asks for that account's password.",
      "depends_on": ["step_reauthenticateM365_selectOriginal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["step_retry_timeout:60"]
    },
    {
      "step_id": "step_reauthenticateM365_password_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": {{json:accountPassword}}
      },
      "description": "Enter the original account password into the Microsoft sign-in password field.",
      "depends_on": [
        "step_reauthenticateM365_assertPassword_{{text:instanceSuffix}}"
      ],
      "preconditions": [
        "dhash:512:384:0:20:1392e8d8d9f6e6e4"
      ],
      "postconditions": [],
      "tags": [
        "precondition_wait_timeout:60"
      ]
    },
    {
      "step_id": "step_reauthenticateM365_submit_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Submit the original Microsoft 365 account password.",
      "depends_on": [
        "step_reauthenticateM365_password_{{text:instanceSuffix}}"
      ],
      "preconditions": [
        "dhash:512:384:0:20:1392e8d8d9f6e6e4"
      ],
      "postconditions": [],
      "tags": []
    },
    {
      "step_id": "step_reauthenticateM365_closeBrowser_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 1006,
        "y": 22
      },
      "description": "Close the browser after the Visual Studio Code sign-in confirmation page appears.",
      "depends_on": [
        "step_reauthenticateM365_submit_{{text:instanceSuffix}}"
      ],
      "preconditions": [
        "dhash:1006:22:16:5:46671d1d674699d7",
        "dhash:1006:22:96:5:926363639200c6c4",
        "dhash:1006:22:0:10:1312094769614541"
      ],
      "postconditions": [],
      "tags": [
        "precondition_wait_timeout:60"
      ]
    }
  ]
}
