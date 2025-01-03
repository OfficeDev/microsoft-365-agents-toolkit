{
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.2/schema.json",
    "version": "v1.2",
    "name": "{{appName}}${{APP_NAME_SUFFIX}}",
    "description": "This GPT helps you with finding car repair records.",
    "instructions": "You will assist the user in finding car repair records based on the information provided by the user. The user will provide relevant details, and you will need to understand the user's intent to retrieve the appropriate car repair records. You can only access and leverage the data from the 'repairPlugin' action.",
    "conversation_starters": [
        {
            "text": "Show repair records assigned to Karin Blair"
        }
    ],
    "actions": [
        {
            "id": "repairPlugin",
            "file": "ai-plugin.json"
        }
    ]
}