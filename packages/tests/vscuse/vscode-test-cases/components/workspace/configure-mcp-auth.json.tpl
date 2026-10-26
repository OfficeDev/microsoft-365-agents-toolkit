{
  "component": {
    "version": 1,
    "id": "configureMcpAuth",
    "parameters": ["instanceSuffix", "scriptBase64", "authMode"]
  },
  "steps": [
    {
      "step_id": "step_configureMcpAuth_openTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+shift+~" },
      "description": "Open a new integrated terminal for the MCP fixture.",
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
      "step_id": "step_configureMcpAuth_assertTerminal_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion a new VS Code integrated terminal is visible and focused with a Bash shell prompt ready for input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_configureMcpAuth_openTerminal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "step_retry_timeout:30"]
    },
    {
      "step_id": "step_configureMcpAuth_typeCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "clear; PROJECT_DIR=\"/home/vscode/AgentsToolkitProjects/${{var:app_name}}\" SCRIPT_B64=\"{{text:scriptBase64}}\" python3 -c 'import base64,os;exec(base64.b64decode(os.environ[\"SCRIPT_B64\"]))'"
      },
      "description": "Type the exact compiler-owned MCP {{text:authMode}} preparation command. Credentials are read from the terminal's runtime environment, never echoed.",
      "content_refs": [],
      "timeout": 60,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_configureMcpAuth_assertTerminal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:workspace",
        "operation:mcp-auth-fixture",
        "mcp_auth_mode:{{text:authMode}}"
      ]
    },
    {
      "step_id": "step_configureMcpAuth_runCommand_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "enter" },
      "description": "Execute the exact MCP preparation command.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_configureMcpAuth_typeCommand_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace"]
    },
    {
      "step_id": "step_configureMcpAuth_assertReady_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the active integrated terminal displays the completed output line VSCUSE_MCP_AUTH_FIXTURE_READY ${{env:MCP_AUTH_RUN_ID}} after execution, with no Python traceback or execution error. A typed command, prior terminal output, or a different run ID is not success.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_configureMcpAuth_runCommand_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "assertion:mcp-fixture-ready", "step_retry_timeout:60"]
    },
    {
      "step_id": "step_configureMcpAuth_closeTerminal_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+`" },
      "description": "Close the verified integrated terminal panel.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_configureMcpAuth_assertReady_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace"]
    },
    {
      "step_id": "step_configureMcpAuth_assertClosed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the integrated terminal panel is no longer visible and the Visual Studio Code workbench is ready.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_configureMcpAuth_closeTerminal_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:workspace", "step_retry_timeout:30"]
    }
  ]
}
