{
  "component": {
    "version": 1,
    "uiSurface": "authentication",
    "account": "m365",
    "id": "signInM365",
    "parameters": ["instanceSuffix", "accountName", "accountPassword"]
  },
  "steps": [
    {
      "step_id": "step_signInM365_assertOption_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the ACCOUNTS section of the side bar lists an entry whose label begins with Sign in to Microsoft.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "entry_state:toolkit-side-bar-visible",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_signInM365_selectOption_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 213,
        "y": 102
      },
      "description": "Click the entry whose label begins with \"Sign in to Microsoft\" in the ACCOUNTS section of the side bar.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_assertOption_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:213:102:16:5:9c63d67674d4c718",
        "dhash:213:102:96:5:8c512758d92c2cce",
        "dhash:213:102:0:10:92c86223a3622421"
      ],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365"
      ]
    },
    {
      "step_id": "step_signInM365_confirmSignIn_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 762,
        "y": 97
      },
      "description": "Click the \"Sign in\" button within the Microsoft 365 developer sandbox modal.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_selectOption_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:762:97:16:5:24b1a72ba9aba343",
        "dhash:762:97:96:5:0008304b0f344900",
        "dhash:762:97:0:10:9c68332223232421"
      ],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365"
      ]
    },
    {
      "step_id": "step_signInM365_assertAccount_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft Sign in page is visible in Chrome with an email or username input and a Next button.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_confirmSignIn_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "account:m365", "step_retry_timeout: 60"]
    },
    {
      "step_id": "step_signInM365_focusAccount_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 369,
        "y": 350
      },
      "description": "Click the email or username input above the Next button on the Microsoft Sign in page in Chrome, not the Use passkey from another device suggestion.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_assertAccount_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "ocr:true"
      ]
    },
    {
      "step_id": "step_signInM365_assertFocusedAccount_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft Sign in page is visible in Chrome with an email or username input that is focused and a Next button; input focus is evidenced by a visible caret, an active underline, or Chrome's \"Use passkey from another device\" suggestion anchored below that input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_focusAccount_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "account:m365", "step_retry_timeout: 60"]
    },
    {
      "step_id": "step_signInM365_typeAccount_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": {{json:accountName}}
      },
      "description": "Enter the resolved Microsoft 365 account name into the email or username input field on the Microsoft Sign in page.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_assertFocusedAccount_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "delay: 3"
      ]
    },
    {
      "step_id": "step_signInM365_assertEnteredAccount_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft Sign in page in Chrome shows {{text:accountName}} in the account input and a Next button.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_typeAccount_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "account:m365", "step_retry_timeout: 60"]
    },
    {
      "step_id": "step_signInM365_next_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 629,
        "y": 484
      },
      "description": "Click the blue Next button on the Microsoft Sign in page in Chrome to continue with the entered account.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_assertEnteredAccount_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "delay:3",
        "ocr:true"
      ]
    },
    {
      "step_id": "step_signInM365_assertPassword_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft Sign in page in Chrome shows {{text:accountName}} and a password input that is visible and focused.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_next_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "account:m365", "step_retry_timeout: 120"]
    },
    {
      "step_id": "step_signInM365_typePassword_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": {{json:accountPassword}}
      },
      "description": "Type the resolved Microsoft 365 account password into the Password field on the Microsoft login page.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_assertPassword_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "delay: 3"
      ]
    },
    {
      "step_id": "step_signInM365_submit_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Press Enter to submit the Microsoft 365 login form from the password entry screen.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_typePassword_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "delay: 3"
      ]
    },
    {
      "step_id": "step_signInM365_closeBrowser_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": {
        "button": "left",
        "x": 1004,
        "y": 19
      },
      "description": "Click the \"Close\" button (red cross icon) in the browser tab bar to close the \"M365 Account - Sign In\" page.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_submit_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:1004:19:16:5:aac833964c9633cc",
        "dhash:1004:19:96:5:d2232323c200e6e6",
        "dhash:1004:19:0:10:0b01410169414141"
      ],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "delay: 3"
      ]
    },
    {
      "step_id": "step_signInM365_assertReady_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the \"ACCOUNTS\" section lists an entry whose label is {{text:accountName}} or that name cut short by the side bar with a trailing ellipsis.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_signInM365_closeBrowser_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "account:m365",
        "readiness:account-visible",
        "step_retry_timeout: 180"
      ]
    }
  ]
}