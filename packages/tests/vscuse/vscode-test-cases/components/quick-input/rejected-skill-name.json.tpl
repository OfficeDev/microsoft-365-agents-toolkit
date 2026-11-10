{
  "component": {
    "version": 1,
    "id": "rejectedSkillName",
    "parameters": ["instanceSuffix", "rejectedName"]
  },
  "steps": [
    {
      "step_id": "step_rejectedSkillName_assertPrompt_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the active text prompt is titled Skill Name.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:quick-input",
        "operation:add-da-skill",
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_rejectedSkillName_type_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "type_text",
      "parameters": { "text": {{json:rejectedName}} },
      "description": "Enter the invalid skill name {{text:rejectedName}}.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_rejectedSkillName_assertPrompt_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:quick-input", "operation:add-da-skill"]
    },
    {
      "step_id": "step_rejectedSkillName_submit_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "enter" },
      "description": "Press Enter to attempt submission of the invalid skill name.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_rejectedSkillName_type_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:quick-input", "operation:add-da-skill"]
    },
    {
      "step_id": "step_rejectedSkillName_assertError_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Skill Name prompt remains open with {{text:rejectedName}} in its input and the validation message \"Skill name may only contain lowercase letters, numbers, and hyphens. Use 1-64 characters, without leading, trailing, or consecutive hyphens.\" The Skill Description prompt is not open.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_rejectedSkillName_submit_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:quick-input",
        "operation:add-da-skill",
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_rejectedSkillName_select_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "keyboard_shortcut",
      "parameters": { "keys": "ctrl+a" },
      "description": "Select all text in the rejected Skill Name input.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_rejectedSkillName_assertError_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:quick-input", "operation:add-da-skill"]
    },
    {
      "step_id": "step_rejectedSkillName_clear_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "key_press",
      "parameters": { "key": "backspace" },
      "description": "Clear the selected invalid skill name before entering its valid replacement.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_rejectedSkillName_select_{{text:instanceSuffix}}"],
      "preconditions": [],
      "postconditions": [],
      "tags": ["component:quick-input", "operation:add-da-skill"]
    }
  ]
}
