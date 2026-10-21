{
  "component": {
    "version": 1,
    "id": "developerPortalSignIn",
    "parameters": ["instanceSuffix", "accountPassword"]
  },
  "steps": [
    {
      "step_id": "step_developerPortalSignIn_assertPassword_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft sign-in password page is visible for the account retained by the preceding Microsoft 365 login, with the password input focused.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "provider:m365",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_developerPortalSignIn_password_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "{{text:accountPassword}}"
      },
      "description": "Type the resolved Microsoft 365 password into the Developer Portal sign-in page.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_developerPortalSignIn_assertPassword_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "provider:m365"]
    },
    {
      "step_id": "step_developerPortalSignIn_submit_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Press Enter to submit the Microsoft 365 password from the focused password input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_developerPortalSignIn_password_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "provider:m365"]
    },
    {
      "step_id": "step_developerPortalSignIn_assertStaySignedIn_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft Stay signed in prompt is visible with Yes focused.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_developerPortalSignIn_submit_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:authentication",
        "provider:m365",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_developerPortalSignIn_staySignedIn_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Press Enter to confirm the focused Yes button on the Stay signed in prompt.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_developerPortalSignIn_assertStaySignedIn_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:authentication", "provider:m365"]
    }
  ]
}
