{
  "component": {
    "version": 1,
    "id": "configurePluginManifest",
    "parameters": ["instanceSuffix", "mutationScriptBase64"]
  },
  "steps": [
    {
      "step_id": "step_configurePluginManifest_open_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+shift+~" },
      "description": "Open a new integrated terminal in Visual Studio Code.",
      "tags": ["component:workspace", "operation:configure-plugin-manifest"]
    },
    {
      "step_id": "step_configurePluginManifest_ready_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion a new VS Code integrated terminal is visible and focused with a Bash shell prompt ready for input.",
      "tags": [
        "component:workspace",
        "operation:configure-plugin-manifest",
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_configurePluginManifest_type_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": {
        "text": "PROJECT_DIR=\"/home/vscode/AgentsToolkitProjects/${{var:app_name}}\" python3 -c 'import base64;exec(base64.b64decode(\"{{text:mutationScriptBase64}}\"))' && printf '\\nVSCUSE_PLUGIN_MANIFEST_%s\\n' CONFIGURED"
      },
      "description": "Type the compiler-owned structured plugin manifest mutation into the active terminal.",
      "tags": ["component:workspace", "operation:configure-plugin-manifest"]
    },
    {
      "step_id": "step_configurePluginManifest_execute_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "enter" },
      "description": "Execute the structured plugin manifest mutation.",
      "tags": ["component:workspace", "operation:configure-plugin-manifest"]
    },
    {
      "step_id": "step_configurePluginManifest_verified_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the integrated terminal displays the complete standalone text VSCUSE_PLUGIN_MANIFEST_CONFIGURED on its own output line, not merely in the typed command, and no Python traceback or assertion error.",
      "tags": [
        "component:workspace",
        "operation:configure-plugin-manifest",
        "step_retry_timeout:120"
      ]
    },
    {
      "step_id": "step_configurePluginManifest_close_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+`" },
      "description": "Close the verified integrated terminal.",
      "tags": ["component:workspace", "operation:configure-plugin-manifest"]
    },
    {
      "step_id": "step_configurePluginManifest_closed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the integrated terminal panel is no longer visible and the Visual Studio Code workbench is ready.",
      "tags": [
        "component:workspace",
        "operation:configure-plugin-manifest",
        "step_retry_timeout:30"
      ]
    }
  ]
}
