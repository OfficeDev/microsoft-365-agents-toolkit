{
  "component": {
    "version": 1,
    "id": "setProjectEnvironmentVariable",
    "parameters": [
      "instanceSuffix",
      "mutationScriptBase64",
      "variableName",
      "variableValue"
    ]
  },
  "steps": [
    {
      "step_id": "step_setProjectEnvironmentVariable_openTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {
        "keys": "ctrl+shift+~"
      },
      "description": "Open a new integrated terminal in Visual Studio Code with Ctrl+Shift+~.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertTerminal_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion a new VS Code integrated terminal is visible and focused with a Bash shell prompt ready for input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_openTerminal_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_typeCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "read -rs VARIABLE_VALUE && PROJECT_DIR=\"/home/vscode/AgentsToolkitProjects/${{var:app_name}}\" VARIABLE_NAME=\"{{text:variableName}}\" VARIABLE_VALUE=\"$VARIABLE_VALUE\" python3 -c 'import base64;exec(base64.b64decode(\"{{text:mutationScriptBase64}}\"))' && printf '\\nVSCUSE_PROJECT_ENVIRONMENT_%s\\n' UPDATED"
      },
      "description": "Type the compiler-owned {{text:variableName}} mutation command for the dev project environment into the active integrated terminal.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertTerminal_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_runCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Press Enter to run the {{text:variableName}} mutation command and begin hidden value input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_typeCommand_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_typeValue_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "{{text:variableValue}}"
      },
      "description": "Type the resolved {{text:variableName}} value into the Bash read command, which does not echo the value.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_runCommand_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_submitValue_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Press Enter to submit the hidden {{text:variableName}} value and finish the verified mutation.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_typeValue_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_beginNotificationsChord_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {
        "keys": "ctrl+k"
      },
      "description": "Press Ctrl+K to begin the Show Notifications keyboard chord.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_submitValue_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_showNotifications_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {
        "keys": "ctrl+shift+n"
      },
      "description": "Press Ctrl+Shift+N to complete the Show Notifications chord and focus Notification Center.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_beginNotificationsChord_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertNotificationsFocused_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Visual Studio Code Notification Center is visible and its notification list has keyboard focus, with the Command Palette closed.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_showNotifications_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_hideNotifications_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "escape"
      },
      "description": "Press Escape while Notification Center is visible and focused to hide it.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertNotificationsFocused_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertNotificationsHidden_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Visual Studio Code Notification Center and Command Palette are no longer visible and the integrated terminal output is unobscured.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_hideNotifications_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_openFocusTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "f1"
      },
      "description": "Press the F1 key to open the Command Palette in Visual Studio Code.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertNotificationsHidden_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertFocusTerminalPalette_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Visual Studio Code Command Palette is visible with a > character in its input box and is ready to accept a command search.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_openFocusTerminal_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_filterFocusTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "Terminal: Focus Terminal"
      },
      "description": "Type Terminal: Focus Terminal into the active Command Palette.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertFocusTerminalPalette_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertFocusTerminalCommand_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Command Palette input box reads >Terminal: Focus Terminal and the highlighted command listed under it is titled Terminal: Focus Terminal.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_filterFocusTerminal_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_executeFocusTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": {
        "key": "enter"
      },
      "description": "Press Enter to focus the existing integrated terminal without creating a new terminal or toggling the panel closed.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertFocusTerminalCommand_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertTerminalFocused_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the existing VS Code integrated terminal is visible, focused, and unobscured, with Notification Center and the Command Palette closed.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_executeFocusTerminal_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertUpdated_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the VS Code integrated terminal visibly shows the exact text VSCUSE_PROJECT_ENVIRONMENT_UPDATED on a standalone line, followed by a Bash shell prompt.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertTerminalFocused_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 120"
      ]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_closeTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {
        "keys": "ctrl+`"
      },
      "description": "Close the verified integrated terminal with Ctrl+`.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_assertUpdated_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "operation:project-environment"]
    },
    {
      "step_id": "step_setProjectEnvironmentVariable_assertClosed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the integrated terminal panel is no longer visible and the Visual Studio Code workbench is ready.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [
        "step_setProjectEnvironmentVariable_closeTerminal_{{text:instanceSuffix}}"
      ],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:project-environment",
        "step_retry_timeout: 30"
      ]
    }
  ]
}
