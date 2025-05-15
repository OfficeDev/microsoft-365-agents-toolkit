{
  "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.3/schema.json",
  "schema_version": "v2.3",
  "name_for_human": "Add-in Skill + Agent for {{appName}}",
  "description_for_human": "Get answer for user's question related to Microsoft 365 products",
  "functions": [
    {
      "name": "addfooter",
      "description": "Action addfooter: take in arg a JSON object, with a footer message in the field 'Footer'.",
      "states": {
        "reasoning": {
          "description": "\n# `addfooter(Footer: str = 'example message to be added to footer') -> str`  Action addfooter: take in arg a JSON object with a string field 'Footer', a footer message.",
          "instructions": "\n- Decide whether to invoke `addfooter(Footer: str = 'example message to be added to footer')`:\n - Check the last user message in the `conversation_memory` and the tool invocation history in the `turn_memory`:\n    - Based on the `result` from `turn_memory`, do I need to return answers, Action addfooter: take in arg a JSON object, with a footer message in the field 'Footer'."
        },
        "responding": {
          "description": "",
          "instructions": "reply"
        }
      }
    },
    {
      "name": "fillcolor",
      "description": "Action fillcolor: take in arg a JSON object, a cell location and a color in hex. Cell location is a single cell.",
      "states": {
        "reasoning": {
          "description": "\n# `fillcolor(Cell: str = 'B7', Color: str = '#30d5c8') -> str`  Action fillcolor: take in arg a JSON object, a cell location and a color in hex. Cell location is a single cell.",
          "instructions": "\n- Decide whether to invoke `fillcolor(Cell: str = 'B7', Color: str = '#30d5c8')`:\n - Check the last user message in the `conversation_memory` and the tool invocation history in the `turn_memory`:\n    - Based on the `result` from `turn_memory`, do I need to return answers, Action fillcolor: take in arg a JSON object, a cell location and a color in hex. Cell location is a single cell."
        },
        "responding": {
          "description": "",
          "instructions": "reply"
        }
      }
    },
    {
      "name": "addtexttoslide",
      "description": "Action addtexttoslide: take in arg a JSON object, a text to be added to a slide.",
      "states": {
        "reasoning": {
          "description": "\n# `addtexttoslide(Text: str = 'hello') -> str` Action addtexttoslide: take in arg a JSON object, a text to be added to a slide.",
          "instructions": "\n- Decide whether to invoke `addtexttoslide(Text: str = 'hello')`:\n - Check the last user message in the `conversation_memory` and the tool invocation history in the `turn_memory`:\n    - Based on the `result` from `turn_memory`, do I need to return answers, Action addtexttoslide: take in arg a JSON object, a text to be added to a slide."
        },
        "responding": {
          "description": "",
          "instructions": "reply"
        }
      }
    }
  ],
  "runtimes": [
    {
      "type": "LocalPlugin",
      "spec": {
        "local_endpoint": "Microsoft.Office.Addin"
      },
      "run_for_functions": ["addfooter", "fillcolor", "addtexttoslide"]
    }
  ]
}
