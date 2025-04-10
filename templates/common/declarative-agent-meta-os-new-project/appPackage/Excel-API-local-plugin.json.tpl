{
  "schema_version": "v2.2",
  "name_for_human": "Excel Add-in + Agent for {{appName}}",
  "description_for_human": "Add-in Actions in Agents for {{appName}}",
  "namespace": "addin_function",
  "functions": [
    {
      "name": "fillcolor",
      "description": "fillcolor changes a single cell location to a specific color.",
      "parameters": {
        "type": "object",
        "properties": {
          "Cell": {
            "type": "string",
            "description": "A cell location in the format of A1, B2, etc.",
            "default": "B2"
          },
          "Color": {
            "type": "string",
            "description": "A color in hex format, e.g. #30d5c8",
            "default": "#30d5c8"
          }
        },
        "required": ["Cell", "Color"]
      },
      "returns": {
        "type": "string",
        "description": "A string indicating the result of the action."
      },
      "states": {
        "reasoning": {
          "description": "`fillcolor` changes the color of a single cell based on the grid location and a color value.",
          "instructions": "The user will pass ask for a color that isn't in the hex format needed in most cases, make sure to convert to the closest approximation in the right format."
        },
        "responding": {
          "description": "`fillcolor` changes the color of a single cell based on the grid location and a color value.",
          "instructions": "If there is no error present, tell the user the cell location and color that was set."
        }
      }
    }
  ],
  "runtimes": [
    {
      "type": "LocalPlugin",
      "spec": {
        "local_endpoint": "ms-office-addin"
      },
      "run_for_functions": ["fillcolor"]
    }
  ]
}
