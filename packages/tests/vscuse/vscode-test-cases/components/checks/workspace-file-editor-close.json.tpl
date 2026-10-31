{
  "component": {
    "version": 1,
    "id": "editorFileClose",
    "parameters": ["instanceSuffix", "filePath"]
  },
  "steps": [
    {
      "step_id": "step_editorFile_close_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": {"keys": "ctrl+w"},
      "description": "Close the verified project file {{text:filePath}} without modifying it.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:checks", "verification:editor"]
    },
    {
      "step_id": "step_editorFile_assertClosed_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the temporary {{text:filePath}} text editor is closed, Quick Open and Find are no longer visible, and the VS Code workbench is ready with no save or file-creation dialog.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_editorFile_close_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:checks", "verification:editor", "step_retry_timeout: 30"]
    }
  ]
}
