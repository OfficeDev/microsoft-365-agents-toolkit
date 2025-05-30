{
  "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.3/schema.json",
  "version": "v1.3",
  "name": "Add-in Skill + Agent for {{appName}}",
  "description": "You are an agent for working with add-in. You can work with any cells, not only well formatted table.",
  "instructions": "You are an agent for working with add-in. You can work with any cells, not only well formatted table.",
  "conversation_starters": [
    {
      "title": "Change cell color (for excel)",
      "text": "Change the cell below A2 to the color of grass. Tell me how long it took in seconds."
    },
    {
      "title": "Add footer (for word)",
      "text": "Add a footer with message 'Hello Agent!'. Tell me how long it took in seconds."
    },
    {
      "title": "Add text to slide (for powerpoint)",
      "text": "Please add text 'Hello PPT!' to the slide. Tell me how long it took in seconds."
    }
  ],
  "actions": [
    {
      "id": "alchemyPlugin",
      "file": "alchemy-plugin.json"
    }
  ]
}
