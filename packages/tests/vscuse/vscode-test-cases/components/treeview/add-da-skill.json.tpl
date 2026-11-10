{
  "component": {
    "version": 1,
    "id": "addDaSkill",
    "uiSurface": "toolkit-development",
    "parameters": ["instanceSuffix"]
  },
  "steps": [
    {
      "step_id": "step_addDaSkill_assertEntry_{{text:instanceSuffix}}",
      "agent": "assertion",
      "tool": "",
      "parameters": {},
      "description": "@assertion the Microsoft 365 Agents Toolkit sidebar has an Add Skill entry under Development, without a Frontier or Preview qualifier.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": [],
      "preconditions": [],
      "postconditions": [],
      "tags": [
        "component:treeview",
        "operation:add-da-skill",
        "step_retry_timeout:30"
      ]
    },
    {
      "step_id": "step_addDaSkill_clickEntry_{{text:instanceSuffix}}",
      "agent": "interaction",
      "tool": "click",
      "parameters": { "button": "left", "x": 110, "y": 490 },
      "description": "Click Add Skill under Development in the Microsoft 365 Agents Toolkit sidebar.",
      "content_refs": [],
      "timeout": 30,
      "retry_count": 0,
      "continue_on_error": "false",
      "depends_on": ["step_addDaSkill_assertEntry_{{text:instanceSuffix}}"],
      "preconditions": [
        "dhash:110:490:16:5:b2c3c644474d0ea9",
        "dhash:110:490:96:5:6a4b676c6823ccd0"
      ],
      "postconditions": [],
      "tags": [
        "component:treeview",
        "operation:add-da-skill",
        "recording:r_1104_061308"
      ]
    }
  ]
}
