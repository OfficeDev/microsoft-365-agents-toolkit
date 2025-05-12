{
  version: "v1.3",
  "name": "Excel Add-in + Agent for {{appName}}",
  "description": "Agent for working with Excel cells.",
  "instructions": "You are an agent for working with an add-in. You can work with any cells, not just a well-formatted table.",
  "conversation_starters": [
    {
      "title": "Change cell color",
      "text": "I want to change the color of cell B2 to orange"
    }
  ],
  "actions": [
    {
      "id": "localExcelPlugin",
      "file": "Excel-API-local-plugin.json"
    }
  ]
}
