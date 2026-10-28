{
  "component": {
    "version": 1,
    "id": "dcrUpdateStage",
    "parameters": ["instanceSuffix", "scriptBase64", "phase"]
  },
  "steps": [
    {
      "step_id": "step_dcrUpdate_openTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+shift+~" },
      "description": "Open a fresh terminal for the DCR update {{text:phase}} stage.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace"]
    },
    {
      "step_id": "step_dcrUpdate_assertTerminal_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion a new VS Code integrated terminal is visible and focused with a Bash shell prompt ready for input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_dcrUpdate_openTerminal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "step_retry_timeout:30"]
    },
    {
      "step_id": "step_dcrUpdate_typeCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "clear; PROJECT_DIR=\"/home/vscode/AgentsToolkitProjects/${{var:app_name}}\" SCRIPT_B64=\"{{text:scriptBase64}}\" python3 -c 'import base64,os;exec(base64.b64decode(os.environ[\"SCRIPT_B64\"]))'"
      },
      "description": "Type the exact DCR update {{text:phase}} program. Helpers prepare files or read the service; only VS Code Provision performs the update.",
      "content_refs": [],
      "timeout": 60,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_dcrUpdate_assertTerminal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:dcr-update-stage",
        "dcr_update_phase:{{text:phase}}",
        "mcp_auth_mode:dcr"
      ]
    },
    {
      "step_id": "step_dcrUpdate_runCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "enter" },
      "description": "Execute the exact DCR update {{text:phase}} stage.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_dcrUpdate_typeCommand_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace"]
    },
    {
      "step_id": "step_dcrUpdate_assertReady_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the active terminal displays a completed output line VSCUSE_DCR_UPDATE_READY ${{var:app_name}} {{text:phase}} followed by the project ownership UUID. No traceback, failed read, authentication error or fixture error is present. A typed command, another phase or project, or prior output is not success.",
      "content_refs": [],
      "timeout": 120,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_dcrUpdate_runCommand_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "dcr_update_ready:{{text:phase}}",
        "step_retry_timeout:90"
      ]
    },
    {
      "step_id": "step_dcrUpdate_closeTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+`" },
      "description": "Close the verified DCR update terminal.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_dcrUpdate_assertReady_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace"]
    },
    {
      "step_id": "step_dcrUpdate_assertClosed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the integrated terminal panel is no longer visible and the Visual Studio Code workbench is ready.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_dcrUpdate_closeTerminal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "step_retry_timeout:30"]
    }
  ]
}
