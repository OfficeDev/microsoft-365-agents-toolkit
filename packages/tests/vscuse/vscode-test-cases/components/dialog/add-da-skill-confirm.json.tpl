{
  "component": {
    "version": 1,
    "id": "confirmAddDaSkill",
    "uiSurface": "dialog",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_confirmAddDaSkill_assertDialog_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion a confirmation dialog says Microsoft 365 Agents Toolkit will modify files in the \"appPackage\" folder to add the agent skill, and its Add button is visible.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:dialog",
        "operation:add-da-skill",
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_confirmAddDaSkill_add_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 487, "y": 98 },
      "description": "Click Add in the confirmation dialog to create and register the skill.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_confirmAddDaSkill_assertDialog_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:487:98:16:5:a74695a5a5555ca3",
        "dhash:487:98:96:5:5824304c0dc24b93"
      ],
      "postconditions": [],
      "tags": [
        "component:dialog",
        "operation:add-da-skill",
        "recording:r_1104_061308"
      ]
    }
  ]
}
