{
  "component": {
    "version": 1,
    "id": "checkCopilotLicense",
    "parameters": ["instanceSuffix", "accountName", "accountPassword"]
  },
  "steps": [
    {
      "step_id": "step_copilotLicense_assertPicker_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Select a walkthrough to open quick pick is visible.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_filter_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "type_text",
      "parameters": {"text": "Build a Declarative Agent"},
      "description": "Filter the walkthrough picker for Build a Declarative Agent."
    },
    {
      "step_id": "step_copilotLicense_assertSelected_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion Build a Declarative Agent is the highlighted selectable walkthrough result.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_open_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "enter"},
      "description": "Open the highlighted Build a Declarative Agent walkthrough."
    },
    {
      "step_id": "step_copilotLicense_assertWalkthrough_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Build a Declarative Agent walkthrough displays Set up your environment.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_expand_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 512, "y": 431},
      "description": "Click Set up your environment in the Build a Declarative Agent walkthrough.",
      "preconditions": ["dhash:512:431:16:5:d2d46667c4d328a0", "dhash:512:431:96:5:0000a4595128d59a", "dhash:512:431:0:10:24b0949490b17075"],
      "tags": ["ocr:true"]
    },
    {
      "step_id": "step_copilotLicense_assertButton_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the expanded Set up your environment section contains a Check Copilot License button.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_checkSignedOut_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 516, "y": 447},
      "description": "Click the Check Copilot License button in Set up your environment.",
      "preconditions": ["dhash:516:447:16:5:2493f3ec6b936ad2", "dhash:516:447:96:5:75622a1d99660000", "dhash:516:447:0:10:24b0949090b17075"],
      "tags": ["ocr:true"]
    },
    {
      "step_id": "step_copilotLicense_openNotificationCommand_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "f1"},
      "description": "Open the command palette to show the persistent notification center."
    },
    {
      "step_id": "step_copilotLicense_assertNotificationPalette_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the VS Code command palette is open with its command input focused.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_filterNotificationCommand_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "type_text", "parameters": {"text": "Notifications: Show Notifications"},
      "description": "Filter for Notifications: Show Notifications."
    },
    {
      "step_id": "step_copilotLicense_assertNotificationCommand_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion Notifications: Show Notifications is the highlighted selectable command palette result.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_showNotifications_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "enter"},
      "description": "Open the notification center using the highlighted command."
    },
    {
      "step_id": "step_copilotLicense_assertNotification_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the selected notification says You need to sign in your Microsoft 365 account, identifies Microsoft 365 Agents Toolkit as its source, and contains a Sign in button.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_focusNotificationSignIn_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "tab"},
      "description": "Focus the Sign in action of the selected Microsoft 365 notification."
    },
    {
      "step_id": "step_copilotLicense_assertNotificationSignIn_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Sign in button inside the You need to sign in your Microsoft 365 account notification has the keyboard focus outline.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_activateNotificationSignIn_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "enter"},
      "description": "Activate the focused Sign in action in the Microsoft 365 notification."
    },
    {
      "step_id": "step_copilotLicense_assertSignIn_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Microsoft 365 account dialog says Microsoft 365 Agents Toolkit needs a Microsoft 365 account with custom app upload permission, and shows Create a Microsoft 365 developer sandbox, Cancel, and Sign in buttons.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_signIn_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 703, "y": 100},
      "description": "Click the rightmost Sign in button in the Microsoft 365 account dialog at the top, beside Cancel and Create a Microsoft 365 developer sandbox.",
      "preconditions": ["dhash:703:100:16:5:0000000000000000", "dhash:703:100:96:5:00000042b2474971", "dhash:703:100:0:10:9cb89590b0717d7f"],
      "tags": ["ocr:true", "precondition_wait_timeout: 1"]
    },
    {
      "step_id": "step_copilotLicense_assertEmail_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Microsoft account sign-in page is open in the browser with a focused email address input and a Next button.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_typeAccount_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "type_text", "parameters": {"text": {{json:accountName}}},
      "description": "Enter the dedicated Copilot-enabled account into the Microsoft sign-in email input.",
      "preconditions": ["dhash:512:384:0:20:1b28f0d9c7e6dae4"],
      "tags": ["delay:5", "precondition_wait_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_next_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 633, "y": 486},
      "description": "Click Next on the Microsoft sign-in page after entering the dedicated account.",
      "preconditions": ["dhash:633:486:16:5:92aa29a3b24db200", "dhash:633:486:96:5:000004b0300c0000", "dhash:633:486:0:10:1b28f0cbc6e6dae4"],
      "tags": ["ocr:true"]
    },
    {
      "step_id": "step_copilotLicense_assertPassword_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Microsoft sign-in password page is open for {{text:accountName}} and the password input is focused.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_typePassword_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "type_text", "parameters": {"text": {{json:accountPassword}}},
      "description": "Enter the protected password into the focused Microsoft password input.",
      "preconditions": ["dhash:512:384:0:20:1b08f0d9d1e6e6e4"],
      "tags": ["delay:3", "precondition_wait_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_submit_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "enter"},
      "description": "Submit the Microsoft password form.",
      "preconditions": ["dhash:512:384:0:20:1b08f0d9d1e6e6e4"]
    },
    {
      "step_id": "step_copilotLicense_assertCallback_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion the Visual Studio Code M365 sign-in callback page displays You are signed in now and can close this page.",
      "tags": ["step_retry_timeout: 120"]
    },
    {
      "step_id": "step_copilotLicense_closeBrowser_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 254, "y": 19},
      "description": "Close the successfully completed M365 sign-in browser tab using its X button.",
      "preconditions": ["dhash:254:19:16:5:01a1424249420221", "dhash:254:19:96:5:9144640100000004", "dhash:254:19:0:10:9391610169414141"],
      "tags": ["delay:3"]
    },
    {
      "step_id": "step_copilotLicense_closeNotifications_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "key_press", "parameters": {"key": "esc"},
      "description": "Close the notification center after returning to VS Code."
    },
    {
      "step_id": "step_copilotLicense_reexpand_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 512, "y": 431},
      "description": "Click Set up your environment in the walkthrough to return from the automatically selected Build a declarative agent step.",
      "preconditions": ["dhash:512:431:16:5:d2d46667c4d328a0", "dhash:512:431:96:5:0000a4595128d59a", "dhash:512:431:0:10:24b0949490b17075"],
      "tags": ["ocr:true"]
    },
    {
      "step_id": "step_copilotLicense_assertReopened_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion Check Copilot License is visible inside the expanded Set up your environment section.",
      "tags": ["step_retry_timeout: 60"]
    },
    {
      "step_id": "step_copilotLicense_checkSignedIn_{{text:instanceSuffix}}",
      "agent": "interaction", "tool": "click", "parameters": {"button": "left", "x": 540, "y": 446},
      "description": "Click Check Copilot License in Set up your environment for the signed-in account.",
      "preconditions": ["dhash:540:446:16:5:e829abaaaa2bd42b", "dhash:540:446:96:5:929a8572629400d2", "dhash:540:446:0:10:23b89490b0717d7f"],
      "tags": ["ocr:true"]
    },
    {
      "step_id": "step_copilotLicense_assertEnabled_{{text:instanceSuffix}}",
      "agent": "assertion", "tool": "", "parameters": {},
      "description": "@assertion Your Microsoft 365 account has Copilot access enabled is visible in the Toolkit Output panel.",
      "tags": ["step_retry_timeout: 120"]
    }
  ]
}